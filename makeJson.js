var xlsx = require('xlsx');
var workbook = xlsx.readFile('./data/#5forthefight.xlsx');
var worksheet = workbook.Sheets[workbook.SheetNames[3]];
var worksheet1 = workbook.Sheets[workbook.SheetNames[0]];

var jsonfile = require('jsonfile')
var fileName = './data/fiveforthefight.json'

var json = {
		edges: [],
		nodes: []
	};
	
var affiliations = [];
var nodeNames = [];

//loop through first column (people that posted) and add their affiliations to the affiliations array
var lastRow = 2;
while(true)
{
	var cell = worksheet1["A" + lastRow];
	if(cell)
	{
		if(affiliations[cell.v] == undefined)
		{
			affiliations[cell.v] = worksheet1["B" + lastRow].v;
		}
		addToJSON(cell.v, true);
	}
	else
	{
		break;
	}
	lastRow++;
}

//loop through tag columns and add those people to the affiliations array
for(var i = 0; i < 5; i++)
{
	var col;
	switch (i)
	{
		case 0:
			col = "G";
			break;
		case 1:
			col = "H";
			break;
		case 2:
			col = "I";
			break;
		case 3:
			col = "J";
			break;
		case 4:
			col = "K";
			break;
		default:
			col = "G";
			break;
	}
	for (var j = 2; j <= lastRow; j++)
	{
		var cell = worksheet1[col + j];
		if(cell)
		{		
			if(affiliations[cell.v] == undefined)
			{
				if(cell.c)
				{
					affiliations[cell.v] = cell.c[0].t;
				}
				else
				{
					affiliations[cell.v] = "None"
				}
			}
			addToJSON(cell.v, false);
			var fromNode = worksheet1["A" + j];
			if(fromNode.v)
			{
				json.edges.push({source:fromNode.v, target: cell.v, id: col + j, type: "arrow", color:"rgb(0,0,0)"});
			}
			else{
				console.log("error on row " + i);
			}
		}
	}
}

// i = 1;
// while(true)
// {
	// var leftCell = worksheet["A" + i];
	// var rightCell = worksheet["B" + i];
	// if(leftCell && rightCell)
	// {
		// var fromNode = leftCell.v;
		// var toNode = rightCell.v;
		// addToJSON(fromNode, true);
		// addToJSON(toNode, false);
		// json.edges.push({source:fromNode, target: toNode, id: "" + i, type: "arrow", color:"rgb(0,0,0)"});
	// }
	// else
	// {
		// break;
	// }
	// i++;
// }

jsonfile.writeFile(fileName, json, function(err)
{
	//console.error(err);
});

function addToJSON(item, isFromNode)
{		
	if (!nodeNames.includes(item))
	{
		nodeNames.push(item);
		var xcoord = 1000 * Math.cos(2 * Math.PI * Math.random());
		var ycoord = 1000 * Math.cos(2 * Math.PI * Math.random());
		var color;
		switch (affiliations[item])
		{
			case "Sigma Chi":
				color = "rgb(0, 157, 220)";
				break;
			case "Alpha Xi Delta":
				color = "rgb(28, 61, 128)";
				break;
			case "Alpha Sigma Alpha":
				color = "rgb(220, 20, 60)";
				break;
			case "Delta Phi Epsilon":
				color = "rgb(255, 214, 74)";
				break;
			case "Zeta Tau Alpha":
				color = "rgb(64, 224, 208)";
				break;
			case "Sigma Sigma Sigma":
				color = "rgb(114, 71, 156)";
				break;
			case "None":
				color = "rgb(100, 100, 100)";
				break;
			default:
				color = "rgb(100, 100, 100)";
				console.log(item);
				console.log("--" + affiliations[item]);
				break;
		}
		var size = isFromNode ? 3 : 2;
		json.nodes.push({label: item, id: item, x: xcoord, y: ycoord, size:size, color: color});
	}
}