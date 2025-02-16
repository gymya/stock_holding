var express = require('express');
const axios = require('axios');
const convertXLSX = require('@/middleware/convertXLSX');

var router = express.Router();

const getVolKData = async (symbol) => {
  console.log(`Fetching ${symbol} data...`);

  try {
    const VolKData = {
      symbol,
      totalOverbuyVolK: 0,
      totalOversellVolK: 0,
      tradeVolumeRate: 0,
      totalDifferenceVolK1D: 0,
      totalDifferenceVolK5D: 0,
      totalDifferenceVolK10D: 0,
      totalDifferenceVolK20D: 0,
    };
    let sumDifferenceVolK = 0;
    const { data } = await axios.get(
      `https://tw.stock.yahoo.com/_td-stock/api/resource/StockServices.brokerTrades;limit=20;sortBy=-date;symbol=${symbol}.TW`
    );
    for (let i = 0; i < data?.list.length; i++) {
      sumDifferenceVolK += data?.list[i]?.totalDifferenceVolK;
      switch (i) {
        case 4:
          VolKData.totalDifferenceVolK5D = sumDifferenceVolK;
          break;
        case 9:
          VolKData.totalDifferenceVolK10D = sumDifferenceVolK;
          break;
        case 19:
          VolKData.totalDifferenceVolK20D = sumDifferenceVolK;
          break;
      }
    }
    VolKData.totalDifferenceVolK1D = data?.list[0]?.totalDifferenceVolK;
    VolKData.totalOverbuyVolK = data?.list[0]?.totalOverbuyVolK;
    VolKData.totalOversellVolK = data?.list[0]?.totalOversellVolK;
    VolKData.tradeVolumeRate = data?.list[0]?.tradeVolumeRate;

    return new Promise((resolve) => {
      resolve(VolKData);
    });
  } catch (error) {
    console.log(error);
    return new Promise((resolve) => {
      resolve({
        symbol,
        totalOverbuyVolK: 'failed',
        totalOversellVolK: 'failed',
        tradeVolumeRate: 'failed',
        totalDifferenceVolK1D: 'failed',
        totalDifferenceVolK5D: 'failed',
        totalDifferenceVolK10D: 'failed',
        totalDifferenceVolK20D: 'failed',
      });
    });
  }
};

const getStockSymbol = async (req, res, next) => {
  try {
    const { data } = await axios.get(
      'https://openapi.twse.com.tw/v1/opendata/t187ap03_L'
    );
    const symbolArray = data.map((stock) => stock.公司代號);

    res.symbolArray = symbolArray;
    next();
  } catch (error) {
    console.log(error);
    res.status(500).send('Error');
  }
};

const getBrokerTrades = async (req, res, next) => {
  res.VolKData = [];
  // forEach is not working async, so we use for loop
  for (let i = 0; i < res.symbolArray.length; i++) {
    const VolKData = await getVolKData(res.symbolArray[i]);
    res.VolKData.push(VolKData);
  }
  next();
};

const createXLSX = async (req, res, next) => {
  res.excelBuffer = convertXLSX(res.VolKData);
  next();
};

const holdingsMiddleware = [getStockSymbol, getBrokerTrades, createXLSX];

router.get('/holdings', holdingsMiddleware, (req, res, next) => {
  const now = new Date();
  const formattedDate = now.toISOString().slice(0, 10).replace(/-/g, '-');

  res.setHeader(
    'Content-Disposition',
    `attachment; filename=${
      formattedDate + encodeURIComponent('主力資料')
    }.xlsx`
  );
  res.send(res.excelBuffer);
});

module.exports = router;
