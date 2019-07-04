/*
*   导出数据到EXCEL(包含图片)
*   exportToExcel(data1,data2,Obj)
*   data1: 数组 表格头数据 如 ['a','b','c']
*   data2: 数组 表格主内容数据
*         - 格式一:数组项为子数组 且与 data1内容一一对应  如 [['a1','a2','a3'],['b1','b2','b3']]
*         - 格式二:数组项为子对象 对象中的键值对中的值和data1 一一对应  如 [{'a1':1,'a2':2,'a3':3},{'b1':1,'b2':2,'b3':3}]
*   Obj: 对象 可设置对应参数
*       - filename : 文件名
*       - sheetName : sheet名
*       - width : 图片单元格宽度
*       - height : 单元格高度
*/ 

var exportToExcel = function (thData, tbData, OptionObj) {
  var re = /http/; // 检测图片地址
  var opt = {
      filename: 'table',
      sheetName: 'table',
      lineHeight:30,
      width:200,
      height:30,
  }
  //处理IE浏览器不兼容Object.assign()
  if (typeof Object.assign != 'function') {
    Object.assign = function(target) {
      if (target == null) {
        throw new TypeError('Cannot convert undefined or null to object');
      }
      target = Object(target);
      for (var index = 1; index < arguments.length; index++) {
        var source = arguments[index];
        if (source != null) {
          for (var key in source) {
            if (Object.prototype.hasOwnProperty.call(source, key)) {
              target[key] = source[key];
            }
          }
        }
      }
      return target;
    };
  }
  //深拷贝数据
  Object.assign(opt, OptionObj);
  //生成表头内容
  var thead = '<thead><tr>'
  for (var i = 0; i < thData.length; i++) {
    thead += '<th>' + thData[i] + '</th>'
  }
  thead += '</tr></thead>' 
  //生成主体内容
  var tbody = '<tbody>'
  for (var j = 0; j < tbData.length; j++) {
    tbody += '<tr style="height:' + opt.height + 'px;">'
    // 获取每一行数据
    var row = tbData[j] 
    if(row instanceof Array){
      for(var k=0;k<row.length;k++){
        if (re.test(row[k])) {
          tbody += '<td style="width:' + opt.width + 'px; text-align: center; vertical-align: middle"><div style="display:inline"><img src=\'' + row[k] + '\' ' + ' ' + 'width=' + '\"' + opt.width + '\"' + ' ></div></td>';
        } else {
          tbody += '<td style="text-align:center">' + row[k] + '</td>'
        }
      } 

    }else{
      for (var key in row) {
        if (re.test(row[key])) {
          tbody += '<td style="width:' + opt.width + 'px; text-align: center; vertical-align: middle"><div style="display:inline"><img src=\'' + row[key] + '\' ' + ' ' + 'width=' + '\"' + opt.width + '\"' + ' ></div></td>';
        } else {
          tbody += '<td style="text-align:center">' + row[key] + '</td>'
        }
      }
    }

    tbody += '</tr>'
  }
  tbody += '</tbody>';

  var table = thead + tbody; 
  //导出为表格
  exportToExcel(table, opt.filename,opt.sheetName)

  function exportToExcel(data, name,sheetName) {
    var uri = 'data:application/vnd.ms-excel;base64,',
        isIE = navigator.appVersion.indexOf("MSIE 10") !== -1 || (navigator.userAgent.indexOf("Trident") !== -1 && navigator.userAgent.indexOf("rv:11") !== -1), // this works with IE10 and IE11 both :)
        template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',
        ctx = {
          worksheet: sheetName,
          table: data,
        }
    if (isIE) {
      if (typeof Blob !== "undefined") {
          //use blobs if we can
          var fullTemplate= format(template, ctx)
          fullTemplate=[fullTemplate]
          var blob1 = new Blob(fullTemplate, { type: "text/html" })
          window.navigator.msSaveBlob(blob1, name+'.xls' )
      } else {
        txtArea1.document.open("text/html", "replace")
        txtArea1.document.write(format(fullTemplate, ctx))
        txtArea1.document.close();
        txtArea1.focus();
        txtArea1.document.execCommand("SaveAs", true, name+'.xls' )
      }
    } else {
      var link = document.createElement('a')
      link.setAttribute('href', uri + base64(format(template, ctx)))
      link.style = "visibility:hidden"
      link.setAttribute('download', name + '.xls')
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link) 
    }
  }
  function base64(s) {
    return window.btoa(unescape(encodeURIComponent(s)))
  }
  function format(s, c){
    return s.replace(/{(\w+)}/g, function (m, p) {
      return c[p]
    })
  }
}