var context;
var Result = 0;

function doWork(){
var fileData = document.getElementById("fileUpload").files[0];
var luList = 'Administrations';
var luList2 = 'Departments';
var luList3 = 'Title names'
var data = readCSV(fileData);											//read csv data
SP.SOD.executeFunc('sp.js', 'SP.ClientContext', setup_context);			//setup context
getLookupIDs(context,luList).then(function(LUitems){					//get lookup IDs than push IDs to data array
	data = pushLookupIDs(LUitems,3,data);
});
	
getLookupIDs(context,luList2).then(function(LUitems){
	data = pushLookupIDs(LUitems,4,data);
	
});
getLookupIDs(context,luList3).then(function(LUitems){
	data = pushLookupIDs(LUitems,5,data);
	
});
getuserIDs(context)
	.then(function(_users){								//get user list than update data array with user IDs
		data = pushIDs(_users,1,data);
		data = pushIDs(_users,2,data);
		data = pushIDs(_users,0,data);
		prepareItem(context,data);											//creating list item
		});
		
setTimeout(function(){ alert("Done! Created " + Result + " items"); }, 1000);
}

function setup_context() {
	context = new SP.ClientContext.get_current();
}

function readCSV(fileData) {
		// function reads csv file and returns data array
		var data_array = [];
        var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.csv|.txt)$/;
        if (regex.test($("#fileUpload").val().toLowerCase())) {
            if (typeof (FileReader) != "undefined") {
                var reader = new FileReader();
                reader.onload = function (e) {
					var rows = e.target.result.split("\n");
					var cells;
                    for (var i = 0; i < rows.length; i++) {
						cells = rows[i].split(",");
						data_array[i] = new Array(cells.length);
						for (var j = 0; j < cells.length; j++) {
						  data_array[i][j] = cells[j].replace(/(\r\n|\n|\r)/gm,"");   //delete return marks
						}

                  }
                }
                reader.readAsText(fileData);

            } else {
                alert("This browser does not support HTML5.");
            }
        } else {
            alert("Please upload a valid CSV file.");
        }
		return data_array;

    }

function getuserIDs(_context){ 
	//function return a list of users including UserName, ID, EMail
	var dfd = new $.Deferred();
	var _users;
	var userInfoList = _context.get_web().get_siteUserInfoList();
	var queryUser  = "dc.gov";
	var query = new SP.CamlQuery();
	var viewXml = "<View> \
                    <Query> \
                       <Where> \
						<Contains><FieldRef Name='EMail' /><Value Type='Text'>" + queryUser + "</Value></Contains> \
                       </Where>  \
                    </Query> \
                  </View>";
	query.set_viewXml(viewXml);
	var items = userInfoList.getItems(query);
	_users = _context.loadQuery(items,'Include(UserName, ID, EMail)');
	_context.executeQueryAsync(IDsReady,onRequestFailed);    
    return dfd.promise();

	function IDsReady(){
	 dfd.resolve(_users);
	}
	function onRequestFailed(sender, args) {
        alert('Error: ' + args.get_message());
		dfd.reject();
    }
}

function pushIDs(_users,column,data_array) {
	  // looping through array to substitute usernames with their ID's using users data from SharePoint
	  // _users is a SharePoint users data collection, column is column in array to substitute, data_aray is array to work
	  var userN = " ";
	  var _user;
	  for (var j = 1; j < data_array.length; j++){
		  _user = data_array[j][column];
		for (var i = 0; i < _users.length; i++){	
			userN = " " + _users[i].get_fieldValues().UserName;
			if (userN.toLocaleLowerCase().includes(_user.toLowerCase())){
				data_array[j][column] = _users[i].get_fieldValues().ID
			}
		}
		
    }
	return data_array;
   }

function prepareItem (context,data)   {
	// going through data, assign variables and call function to create items with variables
	var res;
	for (i = 1; i < data.length; i++){
		v1 = Number(data[i][0]);
		v2 = Number(data[i][1]);
		v3 = Number(data[i][2]);
		v4 = Number(data[i][3]);
		v5 = Number(data[i][4]);
		v6 = Number(data[i][5]);
		res += createListItem(context,v1,v2,v3,v4,v5,v6);
	}
	return res;
}
   
function createListItem(_context, c1,c2,c3,c4,c5,c6) {
	// function recieves context and variables which are used to create list item 
	var _list = _context.get_web().get_lists().getByTitle('Path');     // ***** set list name here  ******
	_context.load(_list);
    var itemInfo = new SP.ListItemCreationInformation();
    var _listItem = _list.addItem(itemInfo);
	
    _listItem.set_item('Director', c3);
    _listItem.set_item('Requester', c1);
	_listItem.set_item('Supervisor', c2);
	_listItem.set_item('Administration', c4);
	_listItem.set_item('Unit', c5);
	_listItem.set_item('PositionTitle', c6);
    _listItem.update();
    _context.load(_listItem);
    _context.executeQueryAsync(_onSucceed, _onFail);

	function _onSucceed() {
		Result += 1;
	}

	function _onFail() {
		alert("Something went wrong with creating new item");
	}	
}	

function getLookupIDs(_context,luList){
	// function recieves context and list name and returns all items from this list
	var dfd = new $.Deferred();
	var _list = _context.get_web().get_lists().getByTitle(luList);
	var query = new SP.CamlQuery();
	var viewXml = "<View> \
                    <Query> \
                       <Where> \
						<Neq><FieldRef Name=Title/><Value Type='Text'>NaN</Value></Neq> \
                       </Where>  \
                    </Query> \
                  </View>";
	query.set_viewXml(viewXml);
	var items = _list.getItems(query);
	var LUitems = _context.loadQuery(items,'Include(Title, Id)');	
	_context.executeQueryAsync(ListReady,onRequestFailed);
	return dfd.promise();
	
	function ListReady(){
	dfd.resolve(LUitems);
	}
	function onRequestFailed(sender, args) {
		dfd.reject();
        alert('Error: ' + args.get_message());
    }
}	

function pushLookupIDs(items,column,data_array) {
		// fuction recieves items array, column number in data and data array
		// it search through data array and substitues lookup filed name with its id.
		var dvalue,luID,luTitle;
		for (var j = 1; j < data_array.length; j++){
		  dvalue = data_array[j][column];
		  for (var i = 0; i < items.length; i++){	
			LUvalue = items[i];
			luID = items[i].get_id();
			luTitle = items[i].get_item('Title');
			if (luTitle.toLocaleLowerCase().includes(dvalue.toLowerCase())){
				//alert(luID+dvalue+luTitle);
				data_array[j][column] = luID;
				}
			}
		
		}
		return data_array;
}
