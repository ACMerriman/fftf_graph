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

for(var i = 0; i < 6; i++)
{
	var col;
	switch (i)
	{
		case 0:
			col = "A";
			break;
		case 1:
			col = "G";
			break;
		case 2:
			col = "H";
			break;
		case 3:
			col = "I";
			break;
		case 4:
			col = "J";
			break;
		case 5:
			col = "K";
			break;
		default:
			col = "A";
			break;
	}
	
}

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
	}
	else
	{
		break;
	}
	lastRow++;
}

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
		}
	}
}



console.log(affiliations);
i = 1;
while(true)
{
	var leftCell = worksheet["A" + i];
	var rightCell = worksheet["B" + i];
	if(leftCell && rightCell)
	{
		var fromNode = leftCell.v;
		var toNode = rightCell.v;
		addToJSON(fromNode);
		addToJSON(toNode);
		json.edges.push({source:fromNode, target: toNode, id: "" + i, type: "arrow", color:"rgb(0,0,0)"});
	}
	else
	{
		break;
	}
	i++;
}

jsonfile.writeFile(fileName, json, function(err)
{
	console.error(err);
});

function addToJSON(item)
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
		json.nodes.push({label: item, id: item, x: xcoord, y: ycoord, size:2, color: color});
	}
}