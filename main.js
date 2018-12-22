
// true: readAsBinaryString ; false: readAsArrayBuffer
var rABS = true;
// index of search query to process
var nextAddress = 0;

var searchQueries = [];
var addresses = ["addresses"];
var geoTimer;

var loading = document.getElementById("loading");
/* transpose an array of arrays , https://gist.github.com/femto113/1784503 */
function transpose(a) {
  return a[0].map(function (_, c) { return a.map(function (r) { return r[c]; }); });
}

function handleFile(e) {
  console.log("handling file");
  loading.style.visibility = "visible";
  download.style.visibility = "hidden";
  /*   loading.innerHTML = ".".repeat(dots % 4);
    dots = dots + 1; */
  /*   console.log("dots: ",dots );
    console.log("inner: ", loading.innerHTML); */
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function (e) {
    console.log("loading file");
    var data = e.target.result;
    if (!rABS) data = new Uint8Array(data);
    console.log("converting data");
    var workbook = XLSX.read(data, { type: rABS ? 'binary' : 'array' });
    console.log("reading data");
    workbook.SheetNames.forEach(function (sheetName) {
      console.log("started converting to row");
      var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
      console.log("finished converting to row");
      XL_row_object.forEach(function (element) {
        var query = element['School'] + ',' + element['District'];
        console.log("pushing query: ",query);
        searchQueries.push(query);
      })

    });
    // avoid query limit, delay Map API calls by 400 milliseconds
    geoTimer = setInterval(theNext, 400);
  };

  function performSearch(addr) {
    loading.innerHTML = "loading :" + Math.round(nextAddress * 100.0 / searchQueries.length) + "%";
    function callback(results, status) {
      // grab formatted address from API call
      if (status == google.maps.places.PlacesServiceStatus.OK) {
        var place = results[0];
        console.log("pushing: ", place);
        addresses.push(place['formatted_address']);
      }
      else {
        addresses.push("Unable to find address.");
        console.log('Failed to processes search ', status);
      }
    }

    var service = new google.maps.places.PlacesService(document.createElement('div'));
    var request = {
      query: searchQueries[addr],
      fields: ['formatted_address'],
    }

    service.findPlaceFromQuery(request, callback);

  }
  function theNext() {
    if (nextAddress < searchQueries.length) {
      performSearch(nextAddress);
      ++nextAddress;
    }
    else {
      searchQueries.unshift("title");
      var resWb = XLSX.utils.book_new();
      resWb.Props = {
        Title: "maptemplate"
      };
      resWb.SheetNames.push('test sheet')
      var wsData = [searchQueries, addresses];
      wsData = transpose(wsData);
      var ws = XLSX.utils.aoa_to_sheet(wsData);
      resWb.Sheets['test sheet'] = ws;
      var wbOut = XLSX.write(resWb, { bookType: 'xlsx', type: 'binary' });

      /* convert binary workbook to octet https://redstapler.co/sheetjs-tutorial-create-xlsx/ */
      function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
      }

      loading.style.visibility = 'hidden';

      // show download button when file is ready
      var btn = document.getElementById('download');
      btn.style.visibility = 'visible';
      btn.addEventListener('click', function () {
        saveAs(new Blob([s2ab(wbOut)], { type: "application/octet-stream" }), 'maptemplate.xlsx');
      })
      clearInterval(geoTimer);
    }

  }

  if (rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}

document.getElementById('upload').addEventListener('change', handleFile, false);
