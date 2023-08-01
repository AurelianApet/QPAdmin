/*
	flot-example.js
*/

$(document).ready(init);


function init() {
    
	// example 1 - basic line graphs
	$.plot($("#flot-example-2"),
	[
			{
				label: "Page views",
				color: "orange",
				shadowSize: 0,
				data: [ [0, 310], [1, 350], [2, 410], [3, 435], [4, 510], [5, 650], [6, 580], [7, 595], [8, 760], [9, 825], [10, 695], [11, 995] ],
				lines: {show: true},
				points: {show: true}
			},
			{
				label: "Visitors",
				color: "#9bc747",
				shadowSize: 0,
				data: [ [0, 110], [1, 250], [2, 220], [3, 290], [4, 310], [5, 420], [6, 390], [7, 415], [8, 470], [9, 525], [10, 495], [11, 595] ],
				lines: {show: true},
				points: {show: true}
			},
			{
				label: "Unique visitors",
				color: "#208ed3",
				shadowSize: 0,
				data: [ [0, 30], [1, 120], [2, 120], [3, 80], [4, 240], [5, 260], [6, 190], [7, 300], [8, 280], [9, 335], [10, 245], [11, 295] ],
				lines: {show: true},
				points: {show: true}	
			}
		],
		{
			xaxis: {
				ticks: [
					[0, "Jan"],
					[1, "Feb"],
					[2, "Mar"],
					[3, "Apr"],
					[4, "May"],
					[5, "Jun"],
					[6, "Jul"],
					[7, "Aug"],
					[8, "Sep"],
					[9, "Oct"],
					[10, "Nov"],
					[11, "Dec"]
				]
			},
			
			grid: {
				borderWidth: 0,
				color: "#aaa",
				clickable: "true"
			}
		}
	);
	$.ajax({
	    url:"/getGraphData.aspx",
	    dataType: 'json',
	    type: 'post',
	    success: function(data){
	        var strDate = data.date;
	        var strLoginCount = data.logincount;
	        var strInoutCount = data.inoutcount;
	        var dates = strDate.split(";");
	        var logincounts = strLoginCount.split(";");
	        var inoutcounts = strInoutCount.split(";");
	        
	        var graph_ticks = [];
	        var graph_datas = [];
	        var graph_inouts = [];
	        
	        for(var i = 0; i < 7; i++){
	            var tmpDate = [];
	            var tmpCount = [];
	            var tmpInout = [];
	            
	            tmpDate.push(i);
	            tmpDate.push(dates[i]);
	            graph_ticks.push(tmpDate);
	            
	            tmpCount.push(i);
	            tmpCount.push(logincounts[i]);
	            graph_datas.push(tmpCount);
	            
	            tmpInout.push(i);
	            tmpInout.push(inoutcounts[i]);
	            graph_inouts.push(tmpInout);
	        }
	        
	        $.plot($("#flot-example-1"),
	            [
			        {
				        label: "Login Counts",
				        color: "orange",
				        shadowSize: 0,
				        data: graph_datas,
				        lines: {show: true},
				        points: {show: true}
			        },
			        {
				        label: "Money Inouts",
				        color: "#9bc747",
				        shadowSize: 0,
				        data: graph_inouts,
				        lines: {show: true},
				        points: {show: true}
			        }
		        ],
		        {
			        xaxis: {
				        ticks: graph_ticks
			        },
        			
			        grid: {
				        borderWidth: 0,
				        color: "#aaa",
				        clickable: "true"
			        }
		        }
	        );
	        
	    },
	    error: function(XMLHttpRequest, textStatus, errorThrown) {
        }
	});
}