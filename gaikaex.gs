function convertTradeHistory() {
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('trade_history');

  //前のデータクリア
  mainSheet.getRange("A6:H10000").clearContent();
  //trade_historyデータ取得
  var data = dataSheet.getRange("A2:L"+dataSheet.getLastRow()).getValues();
  data = data.reverse();
  var r=6;
  var total=0;
  data.forEach(function(elements){
    if(elements[8]!=''){
      mainSheet.getRange(r,1).setValue(elements[7]);  //約定日
      mainSheet.getRange(r,2).setValue(elements[2]);  //通貨ペア
      mainSheet.getRange(r,3).setValue(elements[3]);  //売買
      mainSheet.getRange(r,4).setValue(elements[9]);  //新規価格
      mainSheet.getRange(r,5).setValue(elements[4]);  //約定価格
      mainSheet.getRange(r,6).setValue(elements[5]);  //注文数量
      mainSheet.getRange(r,7).setValue(elements[10]);  //売買損益
      total += Number(elements[10]);
      mainSheet.getRange(r,8).setValue(total);  //損益累計

      r++;
    }
  });

  //グラフの縦軸の上限・下限・目盛り数を計算
  var gdata = mainSheet.getRange("H6:H"+mainSheet.getLastRow()).getValues();
  var max = Math.max.apply(null,gdata);
  var min = Math.min.apply(null,gdata);
  var keta = Math.abs(max).toString().length;
  if(Math.abs(min).toString().length > keta) keta = Math.abs(min).toString().length;
  var order = Math.pow(10, (keta-1));
  var axis_max = Math.ceil(max / order)*order; //グラフ縦軸のMAX
  var axis_mix = Math.floor(min / order)*order; //グラフ縦軸のMIX
  var span = order / 10 * 5;
  var axis_count = (axis_max-axis_mix) / span + 1;

  //グラフ描画
  createChart(axis_max, axis_mix, axis_count);
}

//グラフ描画
function createChart(axis_max, axis_mix, axis_count) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  chart = charts[0];
  if(chart!=undefined) sheet.removeChart(chart);    
  chart = sheet.newChart()
  .asLineChart()
  .addRange(spreadsheet.getRange('H6:H263'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(-1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('legend.position', 'none')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.fontSize', 18)
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('series.0.pointSize', 7)
  .setOption('series.0.lineWidth', 2)
  .setRange(axis_max, axis_mix) //グラフのレンジ(min, max)
  .setOption('vAxis.gridlines.count', axis_count) //目盛数
  .setOption('height', 558)
  .setOption('width', 902)
  .setPosition(4, 9, 25, 17)
  .build();
  sheet.insertChart(chart);
}

