function getSampleSize(Effect,Significance,Conversion) {
    let effect = Effect; // Minimum Detectable Effect
    let significance = Significance; // Statistical Significance
    let conversion = Conversion; // Baseline Conversion Rate

    let c = conversion - (conversion * effect);
    let p = Math.abs(conversion * effect);
    let o = conversion * (1 - conversion) + c * (1 - c);
    let n = 2 * significance * o * Math.log(1 + Math.sqrt(o) / p) / (p * p);
    
    return Math.round(n);
}

function xlfNormSDist(x) {
 
// constants
  var a = 0.2316419;
  var a1 =  0.31938153;
  var a2 = -0.356563782;
  var a3 =  1.781477937;
  var a4 = -1.821255978;
  var a5 =  1.330274429;
 
if(x<0.0)
  return 1-xlfNormSDist(-x);
else
  var k = 1.0 / (1.0 + a * x);
  return 1.0 - Math.exp(-x * x / 2.0)/ Math.sqrt(2 * Math.PI) * k
  * (a1 + k * (a2 + k * (a3 + k * (a4 + k * a5)))) ;
}


/* ================================================== */
/* The Normal distribution probability density function (PDF)
   for the specified mean and specified standard deviation: */
function xlfNormalPDF1a (x, mu, sigma) {
    var num = Math.exp(-Math.pow((x - mu), 2) / (2 * Math.pow(sigma, 2)))
    var denom = sigma * Math.sqrt(2 * Math.PI)
    return num / denom
}


function xlfNormalPDF1b (x, mu, sigma) {
  var num = Math.exp(- 1 / 2 * Math.pow((x - mu) / sigma, 2))
  var denom = sigma * Math.sqrt(2 * Math.PI)
  return num / denom
}


/* ================================================== */
/* The Normal distribution probability density function (PDF)
   for the standard normal distribution: */
function xlfNormalSdistPDF (x) {
    var num = Math.exp(-1 / 2 * x * x )
    var denom = Math.sqrt(2 * Math.PI)
    return num / denom
}


var iframe = document.querySelector('[name="galaxy"]').contentWindow.document;

var pagination = iframe.querySelector(".C_PAGINATION_ROWS_LONG > label").innerText.split(" of ");
if (pagination[0].split("1-")[1] < pagination[1]){
	alert("make sure the page is displaying all the rows - currently dipslaying only "+ pagination[0].split("1-")[1] + 
	" there is however a total of "+pagination[1])}

let Label = first = second = third = fourth = fifth = segment = sessions = "empty"
function Dimensions(Label,first,second,third, fourth,fifth,segment,sessions) {
  this.Label = Label;
  this.First = first;
  this.Second = second;
  this.Third = third;
  this.Fourth = fourth;
  this.Fifth = fifth;
  this.Segment = segment;
  this.Sessions = sessions;
  this.Concat = "";
}

function Results(Label,concat,segment,sessions) {
  this.Label = Label;
  this.Segment = segment;
  this.Sessions = sessions;
  this.Concat = concat;
}

function Results_final(Label,concat,segment,control_users,control_conversion,variant_users,control_conversion,variant_conversion,control_cvr,variant_cvr,Pvalue,SampleSize,Days_run) {
  this.Label = Label;
  this.Concat = concat;
  this.Segment = segment;
  this.control_users = control_users;
  this.variant_users = variant_users;
  this.control_conversion = control_conversion;
  this.variant_conversion = variant_conversion;
  this.control_cvr = (( control_cvr* 100 )).toFixed(2);  
  this.variant_cvr = (( variant_cvr* 100 )).toFixed(2);
  this.Uplift=(( uplift* 100 )).toFixed(2);
  this.ConfLevel = (Pvalue*100).toFixed(2);
  /*this.SampleSizeRequired = SampleSize;*/
  if(isNaN(getSampleSize(uplift, .95, control_cvr))){this.SampleSizeRequired=0} else {this.SampleSizeRequired=getSampleSize(uplift, .95, control_cvr)}
  this.DaysToEnoughData=(SampleSize/variant_users/Days_run);
  if(SampleSize<variant_users){this.EnoughData="Enough data collected"} else{this.EnoughData="need more data"};
  if((Pvalue*100).toFixed(2)>0.95 && SampleSize<variant_users ){this.StatSig="Positive & SS"} else if ((Pvalue*100).toFixed(2)<0.05 & SampleSize<variant_users ){this.StatSig="Negative & SS"} else {this.StatSig="Cant conclude yet"};
  
}
var _table_ = document.createElement('table'),
  _tr_ = document.createElement('tr'),
  _th_ = document.createElement('th'),
  _td_ = document.createElement('td');

// Builds the HTML Table out.
function buildHtmlTable(arr) {
  var table = _table_.cloneNode(false),
    columns = addAllColumnHeaders(arr, table);
  for (var i = 0, maxi = arr.length; i < maxi; ++i) {
    var tr = _tr_.cloneNode(false);
    for (var j = 0, maxj = columns.length; j < maxj; ++j) {
      var td = _td_.cloneNode(false);
      cellValue = arr[i][columns[j]];
      td.appendChild(document.createTextNode(arr[i][columns[j]] || ''));
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
  return table;
}

// Adds a header row to the table and returns the set of columns.
// Need to do union of keys from all records as some records may not contain
// all records
function addAllColumnHeaders(arr, table) {
  var columnSet = [],
    tr = _tr_.cloneNode(false);
  for (var i = 0, l = arr.length; i < l; i++) {
    for (var key in arr[i]) {
      if (arr[i].hasOwnProperty(key) && columnSet.indexOf(key) === -1) {
        columnSet.push(key);
        var th = _th_.cloneNode(false);
        th.appendChild(document.createTextNode(key));
        tr.appendChild(th);
      }
    }
  }
  table.appendChild(tr);
  return columnSet;
}
/**number & name of dimensions**/
var dimensions=[]
var dimensions_array=[]
dimensions=iframe.querySelectorAll('[class^="ID-dimension-data-"]')
for(i=0; i < dimensions.length;i++){
	dimensions_array.push(dimensions[i]['className']);
	console.log(dimensions_array);
}

dimensions=new Set(dimensions_array)
dimensions.size
dimensions_name=iframe.querySelectorAll('[data-guidedhelpid*="data-table-dimension-analytics"]');
let dimensions_name_list=[]
for(t=1;t<dimensions_name.length; t++){
	dimensions_name_list.push(dimensions_name[t].innerText);
	
}

/**number of groups**/
var count_group =[]
count_group_list=iframe.querySelectorAll('[class^="ID-row-"]')

for(i=0; i < count_group_list.length;i++){   	
	count_group.push(iframe.querySelectorAll('[class^="ID-row-"]')[i].getAttribute("class"));    
 }

/**loop through each groups**/
var total_array=[]

var sessions_array=[]
for(i=0; i < count_group.filter(e => !e.includes('ACTION-mouse')).length;i++){
	console.log(i);	
	for(j=0; j < count_group.filter(e => e.includes('ACTION-mouse')).length / count_group.filter(e => !e.includes('ACTION-mouse')).length;j++){	
		/*var sessions*/
		var total=[]
		var read_dimensions=[]
		var dimensions_array_columns=[]
		for (k=0; k < dimensions.size;k++){
		read_dimensions="#ID-rowTable > tbody > tr.ID-row-"+i+"> td.ID-dimension-data-"+k
		dimensions_array_columns.push(iframe.querySelector(read_dimensions).innerText);
		}

		segment="#ID-rowTable > tbody > tr.ID-row-"+i+"-0-"+j+".ACTION-mouse.TARGET-row-"+i+"-0-"+j+ "> td:nth-child(2) > span"	
		sessions="#ID-rowTable > tbody > tr.ID-row-"+i+"-0-"+j+".ACTION-mouse.TARGET-row-"+i+"-0-"+j+" > td.COLUMN-analytics\\.visits"
		Label=dimensions_array_columns[0];
		if(dimensions_array_columns[1] == null){"empty" } else{first=dimensions_array_columns[1]};
		if (dimensions_array_columns[2] == null){"empty" } else{second=dimensions_array_columns[2]};
		if (dimensions_array_columns[3] == null){"empty" } else{third=dimensions_array_columns[3]};
		if (dimensions_array_columns[4] == null){"empty" } else{fourth=dimensions_array_columns[4]};
		if (dimensions_array_columns[5] == null){"empty" } else{fifth=dimensions_array_columns[5]};
		segment=iframe.querySelectorAll(segment)[0].innerText;
		sessions=iframe.querySelector(sessions).innerText.split(" (")[0].replace(",","");
		total_array.push(new Dimensions(Label,first,second,third, fourth,fifth,segment,sessions));
		console.log(new Dimensions(Label,first,second,third, fourth,fifth,segment,sessions));

    }
}

/*read the value from the oject*/
const getValue = (arr, str) => {
  let ref = arr;
  const keys = str.split(".");

  keys.forEach(k => {
    ref = ref[k];
  });

  return ref;
};


/*concact string of selected dimensions and assign to concat*/


let grouping_dimensions=""

let unique_concat=[]
let unique_segment=[]
let unique_variant=[]
let Results_array=[]
var group_dim = function(grouping_dimensions){
	unique_concat=[]
	unique_segment=[]
	unique_variant=[]
for (q=0; q <total_array.length; q++ ){
	var concat_object=[]
	unique_segment.push(getValue(total_array, q+".Segment"));
	unique_variant.push(getValue(total_array, q+".Label"));
	for (w=0; w <grouping_dimensions.length; w++ ){
			concat_object.push(getValue(total_array, q+"."+grouping_dimensions[w]));
			total_array[q].Concat=concat_object.join('_');
			console.log(concat_object.join('_'));
	}
	unique_concat.push(concat_object.join('_'));
}


/*select only unique using Set(), then convert back to array with [...], otherwise this create issue to get length in the loops further down*/
unique_concat=[...new Set(unique_concat)].sort()
unique_segment=[...new Set(unique_segment)].sort()
unique_variant=[...new Set(unique_variant)].sort()
Results_array=[]
/*loop through each group iteration and Variant then sum the groups*/
for (e=0; e <unique_concat.length; e++){
	var results_loop=[]
	for (d=0; d <unique_segment.length; d++){
		for (f=0; f <unique_variant.length; f++){
			results_loop=(total_array.filter(  function(row) {return row.Concat == unique_concat[e] && row.Label == unique_variant[f]  && row.Segment == unique_segment[d];}).reduce(function(sumSoFar, row) { return sumSoFar + parseInt(row.Sessions) }, 0));
			Results_array.push(new Results(unique_variant[f],unique_concat[e],unique_segment[d],results_loop));
		}
	}
}
}
group_dim(grouping_dimensions)

let StartDate=iframe.querySelector('.ID-start').textContent
let Date_start = new Date(StartDate)
let EndDate=iframe.querySelector('.ID-end').textContent
let Date_end = new Date(EndDate)
let Days_run=(Date.parse(Date_end)-Date.parse(Date_start))/1000/60/60/24
let users_segment=unique_segment[0]
let control_label = unique_variant[0]
let Results_array_math=[]

var results_display = function(control_label,users_segment){

/*delete users segments as no need to calculate on it*/
var newunique_segment = unique_segment.filter(m => {
    return !users_segment.includes(m);
});
/*loop through the new array and do the math*/
Results_array_math=[]
for (g=0; g <unique_concat.length; g++){
	for (f=0; f <unique_variant.length; f++){
		var control_users=[]
		var control_conversion=[]
		var variant_users=[]
		var variant_conversion=[]
		var control_cvr=[]
		var variant_cvr=[]
		var control_std=[]
		var variant_std=[]
		var StdError_diff=[]
		var Zscore=[]
		var Pvalue=[]
		for (d=0; d < newunique_segment.length; d++){
				control_users=	   (Results_array.filter(  function(row) {return row.Concat == unique_concat[g] && row.Label == control_label  && row.Segment == users_segment;}).reduce(function(sumSoFar, row) { return sumSoFar + parseInt(row.Sessions) }, 0));
				variant_users=	(	Results_array.filter(  function(row) {return row.Concat == unique_concat[g] && row.Label == unique_variant[f]  && row.Segment == users_segment;}).reduce(function(sumSoFar, row) { return sumSoFar + parseInt(row.Sessions) }, 0));
				control_conversion=(Results_array.filter(  function(row) {return row.Concat == unique_concat[g] && row.Label == control_label  && row.Segment == newunique_segment[d];}).reduce(function(sumSoFar, row) { return sumSoFar + parseInt(row.Sessions) }, 0));
				variant_conversion=(Results_array.filter(  function(row) {return row.Concat == unique_concat[g] && row.Label == unique_variant[f]  && row.Segment == newunique_segment[d];}).reduce(function(sumSoFar, row) { return sumSoFar + parseInt(row.Sessions) }, 0));
				control_cvr=control_conversion/control_users
				variant_cvr=variant_conversion/variant_users
				uplift=(variant_cvr-control_cvr)/control_cvr
				control_std=Math.pow((control_cvr*(1-control_cvr)/control_users),(0.5))
				variant_std=Math.pow((variant_cvr*(1-variant_cvr)/variant_users),(0.5))
				StdError_diff=Math.pow((Math.pow(control_std,2)+Math.pow(variant_std,2)),1/2)
				Zscore=(variant_cvr-control_cvr)/StdError_diff
				Pvalue=xlfNormSDist(Zscore)
				SampleSize=getSampleSize(uplift, .95, control_cvr)
				Results_array_math.push(new Results_final(unique_variant[f],unique_concat[g],newunique_segment[d],control_users,control_conversion,variant_users,control_conversion,variant_conversion,control_cvr,variant_cvr,Pvalue,SampleSize,Days_run));
			}
		}
}


}



/*iframe.querySelector('.T_TAB_CONTAINER').innerHTML =table_html */




const newNode = iframe.createElement("div");
newNode.setAttribute("id", "AB_Div");

const RadioNode = iframe.createElement("div");
RadioNode.setAttribute("id", "Radio_Label");

const CheckboxNode = iframe.createElement("div");
CheckboxNode.setAttribute("id", "Checkbox_Label");

const SegmentRadioNode = iframe.createElement("div");
SegmentRadioNode.setAttribute("id", "SegmentRadio_Label");



const btn = iframe.createElement("button");
btn.setAttribute("class", "btn btn-primary");
btn.innerHTML = "Delete All tables";
btn.name = "deleteTable";
btn.onclick=function(){
	var elems = iframe.querySelectorAll(".AB_Results_Table");
	[].forEach.call(elems, function(el) {
	  el.remove();
	});
	}
	
	
const list = iframe.getElementById("ID-tabControl");
list.insertBefore(newNode, list.children[0]);

newNode.appendChild(btn);
newNode.appendChild(RadioNode);
newNode.appendChild(CheckboxNode);
newNode.appendChild(SegmentRadioNode);


/*inject radio and checkbox*/
var output="";
var output2="";
var output3="";
var unique_dimensions=["First","Second","Third","Fourth","Fifth"]
for(var i=0; i< unique_variant.length; i++){
	if( i == 0) {output2+= "Choose the control: "+unique_variant[i]+':<input type="radio" checked="true" value="'+unique_variant[i]+'" name="Label">';	} else {
	output2+= unique_variant[i]+':<input type="radio" value="'+unique_variant[i]+'" name="Label">';	}
	iframe.getElementById("Radio_Label").innerHTML=output2;
}

for(var ii=0; ii< unique_segment.length; ii++){
	if( ii == 0) {output3+="Choose the User base: "  +unique_segment[ii]+':<input type="radio" checked="true" value="'+unique_segment[ii]+'" name="Segment">';	} else {
	output3+= unique_segment[ii]+':<input type="radio" value="'+unique_segment[ii]+'" name="Segment">';	}
	iframe.getElementById("SegmentRadio_Label").innerHTML=output3;
}
output="Split by: "
for(var iii=0; iii< dimensions_name_list.length; iii++){
	output+= '<input type="checkbox" value='+unique_dimensions[iii]+' name="checkbox_dimensions">'+ '   ' + dimensions_name_list[iii]+'   ';
	iframe.getElementById("Checkbox_Label").innerHTML=output;
}

let renderTable =function(Results_array_math){
		
var table_html =document.body.appendChild(buildHtmlTable( Results_array_math));
table_html.setAttribute("class", "AB_Results_Table");
table_html.setAttribute("id", "table"+unique_concat);
table_html.setAttribute("border", "1px solid");
table_html.setAttribute("width", "100%");
table_html.style.color = 'black';
newNode.appendChild(table_html);


/*download csv*/
let csv = '';
let header = Object.keys(Results_array_math[0]).join(',');
let values = Results_array_math.map(o => Object.values(o).join(',')).join('\n');

csv += "data:text/csv;charset=utf-8," + header + '\n' + values;

var encodedUri = encodeURI(csv);
const btnCSV = iframe.createElement("a");
btnCSV.setAttribute("href",encodedUri);
btnCSV.setAttribute("download", StartDate+"_"+EndDate+"("+Days_run+"days_run) "+unique_concat+".csv");
btnCSV.setAttribute("class", "AB_Results_Table");
btnCSV.innerHTML = "Download CSV";
btnCSV.name = "Download CSV";
newNode.appendChild(btnCSV);

}




results_display(control_label,users_segment)
renderTable(Results_array_math)


let grouping_dimensions_compare=[]
let checkboxValues =[]
function timeout() {
	
    setTimeout(function () {
		/*check Label value*/
		control_label_compare = iframe.querySelector('input[name="Label"]:checked').value;
		/*check Segemnt value*/
		segment_compare = iframe.querySelector('input[name="Segment"]:checked').value;
		/*check dimenstions checked*/
		checkboxValues= iframe.querySelectorAll('input[name="checkbox_dimensions"]:checked');
		if (checkboxValues.length == 0) {
			console.log("pas de loop grouping_dimensions_compare"+grouping_dimensions_compare);
			grouping_dimensions_compare=""	
			console.log("assigned grouping_dimensions_compare"+grouping_dimensions_compare);
		} else {
			grouping_dimensions_compare=[]
			for (r=0; r < checkboxValues.length ; r++){				
				grouping_dimensions_compare.push(String(checkboxValues[r].value))

			}
		}
        if ( control_label_compare !=control_label || segment_compare!=users_segment || String(grouping_dimensions_compare) !=String(grouping_dimensions) ){
			group_dim(grouping_dimensions_compare)
			results_display(control_label_compare,segment_compare)
			renderTable(Results_array_math)
			control_label=control_label_compare
			grouping_dimensions=grouping_dimensions_compare
			users_segment=segment_compare			

		}
        // create a recursive loop.
        timeout();
    }, 2500);
}

timeout()
 
 
