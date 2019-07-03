var exportToExcel = function (thData, tbData, OptionObj) {
  var re = /http/, // 检测图片地址
  		opt;
  var option = function(){
    this.name="table"
    this.filename="table"
		this.width= 200
		this.height=30
  }

  typeof OptionObj==='object'?opt = OptionObj :opt = new option()
  console.log(opt.width)
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
    for (var key in row) {
      if (re.test(row[key])) {
        tbody += '<td style="width:' + opt.width + 'px; text-align: center; vertical-align: middle"><div style="display:inline"><img src=\'' + row[key] + '\' ' + ' ' + 'width=' + '\"' + opt.width + '\"' + ' ></div></td>';
      } else {
        tbody += '<td style="text-align:center">' + row[key] + '</td>'
      }
    }
    tbody += '</tr>'
  }
  tbody += '</tbody>';

  var table = thead + tbody; 
  //导出为表格
  exportToExcel(table, opt.filename)

  function exportToExcel(data, name) {
    var uri = 'data:application/vnd.ms-excel;base64,',
				isIE = navigator.appVersion.indexOf("MSIE 10") !== -1 || (navigator.userAgent.indexOf("Trident") !== -1 && navigator.userAgent.indexOf("rv:11") !== -1), // this works with IE10 and IE11 both :)
  			template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',
	  		ctx = {
			    worksheet: name,
			    table: data
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