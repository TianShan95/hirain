console.log('20210515_修改')
//放置三个图例标签
d3.select('#lengendsvg')
.append('line').attr('x1','1200').attr('y1','20').attr('x2','1400').attr('y2','20').attr('stroke','green').attr('stroke-width','3')
d3.select('#lengendsvg').append("text")
    .attr("x", '1410')
    .attr("y", '20')
    .attr("dy", ".35em")
    .attr("fill", "green")
    .text('load_average');

d3.select('#lengendsvg')
.append('line').attr('x1','1200').attr('y1','40').attr('x2','1400').attr('y2','40').attr('stroke','red').attr('stroke-width','3')
d3.select('#lengendsvg').append("text")
    .attr("x", '1410')
    .attr("y", '40')
    .attr("dy", ".35em")
    .attr("fill", "red")
    .text('cpus%_us+sy');

d3.select('#lengendsvg')
.append('line').attr('x1','1200').attr('y1','60').attr('x2','1400').attr('y2','60').attr('stroke','purple').attr('stroke-width','3')
d3.select('#lengendsvg').append("text")
    .attr("x", '1410')
    .attr("y", '60')
    .attr("dy", ".35em")
    .attr("fill", "purple")
    .text('mem_avail');
// get main SVG and its attributes & setting hyper-parameters; 
const svg = d3.select('#mainsvg');
const width = +svg.attr('width');
const height = +svg.attr('height');
const margin = {top: 10, right: 620, bottom: 50, left: 220};
const innerWidth = width - margin.left - margin.right;
const innerHeight = height - margin.top - margin.bottom;
const g = svg.append('g').attr('id', 'maingroup')
.attr("x", '1410')
.attr("y", '40')
.attr('transform', `translate(${margin.left}, ${margin.top})`);
const xValue = (datum) => {return datum['time_1']};
const yValue_01 = (datum) => {return datum['load_average']};
const yValue_02 = (datum) => {return datum['cpus%_us+sy']};
const yValue_03 = (datum) => {return datum['mem_avail']};
let xScale, xScale2, xScale0, yScale_01, yScale_02, yScale_03;
let alldates;
let sequantial;
let aduration = 0;
var dateParse =d3.timeParse("%d-%H:%M:%S");  //将字符串解析为d3的日期对象
var dayParse =d3.timeFormat('%d');  //将d3的时间对象解析到天的函数
var YMDParse =d3.timeFormat('%Y%B%d');  //将d3的时间对象解析到字符串的函数
let c = 0;
let click = 1;
var line_01, line_02, line_03, dots_01, dots_02, dots_03;
var xAxis, xAxisGroup, yAxis_01, yAxisGroup_01, yAxis_02, yAxisGroup_02, yAxis_02, yAxisGroup_02;
let circleupdates_01, circleupdates_02, circleupdates_03, circleenter_01, circleenter_02,circleenter_03;
var ChangeDay = 1, whichDay = 0;  //改变比例尺到哪一天
let record_date=[];
var updataXAxis;
var text_build;  // 显示build信息
var go = null;
// setting up the tip tool; 
const tip_01 = d3.tip()
.attr('class', 'd3-tip').html(function(d) { return 'time: '+(d['time'])+'<br/>'+'load_average: '+d['load_average'] });
svg.call(tip_01);

const tip_02 = d3.tip()
.attr('class', 'd3-tip').html(function(d) { return 'time: '+(d['time'])+'<br/>'+'cpus%_us+sy: '+d['cpus%_us+sy'] });
svg.call(tip_02);

const tip_03 = d3.tip()
.attr('class', 'd3-tip').html(function(d) { return 'time: '+(d['time'])+'<br/>'+'mem_avail: '+d['mem_avail'] });
svg.call(tip_03);

let str;
const tip_04 = d3.tip()
.attr('class', 'd3-tip').html(function(d) { str = d['app'].replace(/\n/g,"<br/>");str1 = str.replace(/\s/g, "&nbsp"); return str1 });
svg.call(tip_04);

function LastBtn(){
  aduration = 2000;
  if(c >= 1){
    c = c - 1;
    click = 1;
    whichDay = 0;
    ChangeDay = 1;
    record_date=[];
  }
  else{
    c = alldates.length - 1;
    click = 1;
    whichDay = 0;
    ChangeDay = 1;
    record_date=[];
  }
};
function NextBtn(){
  aduration = 2000;
  if(c < (alldates.length-1)){
    c = c + 1;
    click = 1;
    whichDay = 0;
    ChangeDay = 1;
    record_date=[];
      }
    else{
      c = 0;
      click = 1;
      whichDay = 0;
      ChangeDay = 1;
      record_date=[];
    }
};
function RestartBtn(){
  aduration = 2000;
  xScale = xScale0;
    xAxis.scale(xScale);
    xAxisGroup.call(xAxis);

    click = 1;
    g.select('#x').selectAll('.tick text')
    .attr('transform', 'translate(0 20) rotate(20)');

};
function LastDay(){
    ChangeDay = -1;  // 上一天数据的标志
    click = 1;
};
function NextDay(){
    ChangeDay = 1;  // 下一天数据的标志
    click = 1;
};

function GoBtn(){  // 跳转函数
  // alert(document.getElementById("myNumber").value);
  aduration = 2000;
  go = document.getElementById("myNumber").value;
  click = 1;
  // console.log(go)
  // console.log("-----------")
};
const render_init = function(data){
xScale = d3.scaleTime()
.domain([dateParse('1-00:00:00'), dateParse('1-23:59:59')])
.range([0, innerWidth])
.nice();

xScale2 = xScale.copy(); // reference.
xScale0 = xScale.copy(); // reference.


yScale_01 = d3.scaleLinear()
.domain([d3.max(data, yValue_01), d3.min(data, yValue_01)])
// .domain(d3.extent(data, yValue_01))
.range([0, innerHeight])
.nice();

yScale_02 = d3.scaleLinear()
.domain([d3.max(data, yValue_02), d3.min(data, yValue_02)])
.range([0, innerHeight])
.nice();

yScale_03 = d3.scaleLinear()
.domain([d3.max(data, yValue_03), d3.min(data, yValue_03)])
.range([0, innerHeight])
.nice();

// Adding axes
xAxis = d3.axisBottom(xScale)
// .ticks(Math.floor(8))
.tickFormat(d3.timeFormat('%H:%M:%S'))
.tickSize(-innerHeight);


xAxisGroup = g.append('g')
.call(xAxis)
.attr('id','x')
.attr("class", "x axis")
.attr('transform', `translate(0, ${innerHeight})`);

yAxis_01 = d3.axisLeft(yScale_01).tickSize(-innerWidth);


yAxisGroup_01 = g.append('g')
.call(yAxis_01)
.attr('id','y_01')
.attr("class", "y axis");

yAxis_02 = d3.axisLeft(yScale_02).tickSize(0);
yAxisGroup_02 = g.append('g').call(yAxis_02).attr('id','y_02')
.attr('transform', `translate(${-30}, 0)`);

yAxisGroup_02.select('path.domain').attr('transform', `translate(${30}, 0)`);

yAxis_03 = d3.axisRight(yScale_03).tickSize(0);
yAxisGroup_03 = g.append('g')
.call(yAxis_03)
.attr('id','y_03')
.attr('transform', `translate(${innerWidth}, 0)`);

d3.select('#lengendsvg').append("text")
        .data(['seq'])
        .attr("x", 780)
        .attr("y", 60)
        .attr("dy", ".5em")
        .style("text-anchor", "end")
        .attr("fill", "#504f4f")
        .attr('font-size', '1.5em')
        .attr('font-weight', 'bold')
        .text(`共${alldates.length}次启动数据`);

g.append('path').attr('id', 'alterPath_01');
g.append('path').attr('id', 'alterPath_02');
g.append('path').attr('id', 'alterPath_03');

g.append('circle').attr('class', 'circle_01');
g.append('circle').attr('id', 'circle_011');

svg.append("rect")
.attr("x", 1550)
.attr("y", 180)
.attr("width",600)
.attr("height",150)
.attr("fill","black")
.attr("opacity", 0.8)

text_build = svg.append("text")
.attr("x",1620).attr("y", 180).attr("fill", "white")



};

var zoom = d3.zoom()
.scaleExtent( [ 0.001, 10000 ] )
.on( "zoom", zoomed );

svg.call(zoom);




function zoomed(){

tip_01.hide();
tip_02.hide();
tip_03.hide();
tip_04.hide();

    xScale = d3.event.transform.rescaleX(xScale2)
    xAxis.scale(xScale);
    xAxisGroup.call(xAxis);

    g.select('#x').selectAll('.tick text')
    .attr('transform', 'translate(0 20) rotate(20)');

    click = 1;
};


svg.append("clipPath")
    .attr("id", "clip")
    .append("rect")
    .attr("width", innerWidth)
    .attr("height", height);

    const isSameDay = function(startTime, endTime) {
      const startTimeMs = new Date(startTime).setHours(0,0,0,0);
      const endTimeMs = new Date(endTime).setHours(0,0,0,0);
      return startTimeMs === endTimeMs ? true : false
    }  // 判断两个日期是否是同一天

    function deepClone(obj){
      let _obj = JSON.stringify(obj),
          objClone = JSON.parse(_obj);
      return objClone
  }  //  普通的深拷贝

  Date.prototype.clone=function(){
    return new Date(this.valueOf());
  }  // 日期的深拷贝

  function unique(arr){
  // 遍历arr，把元素分别放入tmp数组(不存在才放)
  var tmp = new Array();
  for(var i in arr){
    //该元素在tmp内部不存在才允许追加
    if(tmp.indexOf(arr[i])==-1){
      tmp.push(arr[i]);
      console.log('添加')
    }
  }
  return tmp;
  }  // 数组去重


 updataXAxis = function(data){

  let record_whichDay
  let xmin
  
    if(ChangeDay == 1){
        record_whichDay = whichDay;
        whichDay = whichDay + 1;  //如果点击的是下一天数据 则变量加一
    }
    if(ChangeDay == -1 && whichDay > 1){
        record_whichDay = whichDay;
        whichDay = whichDay - 1;  //如果点击的是上一天数据 则变量减一
    }
    let record_day = []  //记录应该在坐标轴显示
    date = Array.from( new Set(data.map( d => d['time_1']) ));   // 取出本次开机数据的时间对象
    console.log('本次开机的所有日期对象');
    console.log(date);

   
   
      /*******如果日期记录没有的话，则将日期去重*********/
    if(record_date.length == 0){
      /************去除重复的日期******************/
      record_date.push(date[0])
      let repeat = 0;
      for(let i=1; i<date.length; i++){
        for(let j=0; j<record_date.length; j++){           
          if(isSameDay(record_date[j],date[i])){
            repeat = 1;
            continue;
          }
        }
        if(repeat == 0){
          record_date.push(date[i]);
        }
        repeat = 0;
      } // 共有record_date.length这么多天的数据
      console.log('日期去重之后的数组: record_data');
      console.log(record_date);
    }

    date.forEach(d => {if(isSameDay(d,record_date[whichDay-1])){record_day.push(d);}});  //把这一天的所有的时间对象存于列表
    console.log('本天的时间数据')
    console.log(record_day)
    if(record_day.length > 0){
      if(record_day.length == 1){
        xmin = dateParse(+whichDay+'-00:00:00');
        console.log(xmin)
        console.log(d3.max(record_day))
      }
      else{
        xmin = d3.min(record_day);
      }
      console.log(isSameDay(xmin,d3.max(record_day))+'fffff')
      xScale = d3.scaleTime()
      .domain([xmin, d3.max(record_day)])
      .range([0, innerWidth-0])
      .nice();
      xAxis.scale(xScale);
      xAxisGroup.call(xAxis);

      xScale2 = xScale.copy(); // reference.
      xScale0 = xScale.copy(); // reference.

      g.select('#x').selectAll('.tick text')
      .attr('transform', 'translate(0 20) rotate(20)');
      console.log(record_day)
      record_day = []
    }
    else{
      whichDay = record_whichDay;
    }
    // console.log('显示哪一天 whichDay: ',whichDay)
    // console.log('上一天还是下一天 ChangeDay: ',ChangeDay)

    d3.select('#lengendsvg').selectAll('.date_text1').remove()
    d3.select('#lengendsvg').append("text")
    .data(['seq']) 
    .attr('class', 'date_text1')
    .attr("x", 450)
    .attr("y", 60)
    .attr("dy", ".5em")
    .style("text-anchor", "end")
    .attr("fill", "#504f4f")
    .attr('font-size', '1.5em')
    .attr('font-weight', 'bold')
    .text(`第${whichDay}段`+`-共${record_date.length}段数据`);

}

const convert_build_info = function(d){

  let str_build = d['version'].replace(/\n/g,"<br/>");
  let str2_build = str_build.split("<br/>");
  console.log("5月8日")
  console.log(str2_build)
  return str2_build
}

const render_update_alter = function(data){
g.selectAll('.circle_01').remove()
circleupdates_01 = g.selectAll('.circle_01')
                    .data(data, d => d['load_average'])
                    .attr('id','circle_011');
circleupdates_02 = g.selectAll('.circle_01').data(data, d => d['cpus%_us+sy'])
                    .attr('id','circle_012');
circleupdates_03 = g.selectAll('.circle_01').data(data, d => d['mem_avail'])
                    .attr('id','circle_013');


circleenter_01 = circleupdates_01.enter().append('circle')
.attr('cx', d => xScale(xValue(d)))
.attr('cy', d => yScale_01(yValue_01(d)))
.attr('r', 3)
.attr('fill', 'green')
.attr("stroke-width",1)
.attr('opacity', 0.8)
.attr('class', 'circle_01')
.attr('id', 'circle_011')
.attr("clip-path", "url(#clip)")
.on('mouseover', function(d){
    tip_01.show(d)
    tip_04.hide(d)
    d3.select(this)
    .attr("opacity", 0.5)
    .attr('r', 10)
    .attr("stroke","white")
    .attr("stroke-width", 3);
})
.on('mouseout', function(d){
    tip_01.hide(d)
    tip_04.hide(d)
      d3.select(this)
      .attr("opacity", 1)
      .attr('r', 3)
      .attr("stroke","green")
      .attr("stroke-width",1);
    })
.on('click', function(d){
tip_01.hide(d)
tip_04.show(d)
});
circleenter_02 = circleupdates_02.enter().append('circle')
.attr('cx', d => xScale(xValue(d)))
.attr('cy', d => yScale_02(yValue_02(d)))
.attr('r', 3)
.attr('fill', 'red')
.attr('opacity', 0.8)
.attr("stroke-width",1)
.attr('class', 'circle_01')
.attr("clip-path", "url(#clip)")
.on('mouseover', function(d){  
  tip_02.show(d)
  tip_04.hide(d)
    d3.select(this)
    
    .attr("opacity", 0.5)
    .attr('r', 10)
    .attr("stroke","white")
    .attr("stroke-width", 3);
})
.on('mouseout', function(d){
  tip_02.hide(d)
  tip_04.hide(d)
      d3.select(this)
      .attr("opacity", 1)
      .attr('r', 3)
      .attr("stroke","red")
      .attr("stroke-width",1);
    })
.on('click', function(d){
tip_02.hide(d)
tip_04.show(d)
});

circleenter_03 = circleupdates_03.enter().append('circle')
.attr('cx', d => xScale(xValue(d)))
.attr('cy', d => yScale_03(yValue_03(d)))
.attr('r', 3)
.attr('stroke', 'purple')
.attr('opacity', 0.8)
.attr('class', 'circle_01')
.attr("clip-path", "url(#clip)")
.on('mouseover', function(d){
  tip_03.show(d)
    d3.select(this)
    .attr("opacity", 0.5)
    .attr('r', 10)
    .attr("stroke","white")
    .attr("stroke-width", 3);
})
.on('mouseout', function(d){
  tip_03.hide(d)
  tip_04.hide(d)
      d3.select(this)
      .attr("opacity", 1)
      .attr('r', 3)
      .attr("stroke","purple")
      .attr("stroke-width",1);
    })
.on('click', function(d){
    tip_03.hide(d)
  tip_04.show(d)
});


// circleupdates_01.merge(circleenter_01)
// .transition().ease(d3.easeLinear).duration(aduration)
// .attr('cx', d => xScale(xValue(d)))
// .attr('cy', d => yScale_01(yValue_01(d)))
// .attr("clip-path", "url(#clip)");


  line_01 = d3.line()
  .x(d => {return xScale(xValue(d))})
  .y(d => {return yScale_01(yValue_01(d))})
//   .curve(d3.curveBasis)
//   .curve(d3.curveCardinal.tension(0.5));


  line_02 = d3.line()
  .x(d => {return xScale(xValue(d))})
  .y(d => {return yScale_02(yValue_02(d))})
  //.curve(d3.curveBasis)
//   .curve(d3.curveCardinal.tension(0.5));

  line_03 = d3.line()
  .x(d => {return xScale(xValue(d))})
  .y(d => {return yScale_03(yValue_03(d))})
  //.curve(d3.curveBasis)
//   .curve(d3.curveCardinal.tension(0.5));


  // See https://github.com/d3/d3-shape/blob/v1.3.7/README.md#curves
  d3.select('#alterPath_01').datum(data)
  .attr('class', 'datacurve_01')
  .attr("fill", "none")
  .attr("stroke", "green")
  .attr("stroke-width", 2.5)
  .transition().duration(aduration)
  .attr("d", line_01)
  .attr("clip-path", "url(#clip)");

  d3.select('#alterPath_02').datum(data)
  .attr('class', 'datacurve_02')
  .attr("fill", "none")
  .attr("stroke", "red")
  .attr("stroke-width", 2.5)
  .transition().duration(aduration)
  .attr("d", line_02)
  .attr("clip-path", "url(#clip)");

  d3.select('#alterPath_03').datum(data)
  .attr('class', 'datacurve_03')
  .attr("fill", "none")
  .attr("stroke", "purple")
  .attr("stroke-width", 2.5)
  .transition().duration(aduration)
  .attr("d", line_03)
  .attr("clip-path", "url(#clip)");

  time = data[0]['boot_times'];



  d3.select('#lengendsvg').selectAll('.date_text').remove()
  d3.select('#lengendsvg').append("text")
    .data(['seq']) 
    .attr('class', 'date_text')
    .attr("x", 450)
    .attr("y", 60)
    .attr("dy", ".5em")
    .style("text-anchor", "left")
    .attr("fill", "#504f4f")
    .attr('font-size', '1.5em')
    .attr('font-weight', 'bold')
    .text(`-第${time}次启动`);

    
   build_info = convert_build_info(data[0]) 
  // d3.select('#lengendsvg').selectAll('.date_text').remove()
  // d3.select('#lengendsvg')
  // .data(build_info)
  //   .enter() 
  console.log("++++++++++++++++++++++")
  console.log(data[0])
  console.log("++++++++++++++++++++")

  text_build.selectAll('.build_info').remove()
  text_build.selectAll("tspan")
            .data(build_info)
            .enter()
            .append("tspan")
            .attr("x", text_build.attr("x"))
            .attr("dy","1.5em")
            .attr("class", "build_info")
            .html(function(d){
              return d;})
    

}


d3.csv('static/data/top.csv').then( data => {
  data.forEach( d => {
      d['boot_times'] = +(d['boot_times']);
      d['top_times'] = +(d['top_times']);
      d['time_1'] = dateParse(d['time']);
      d['load_average'] = +(d['load_average']);
      d['cpus%_us+sy'] = +(d['cpus%_us+sy']);
      d['mem_avail'] = +(d['mem_avail']);
      
  } );
  console.log('全部数据')
  console.log(data);
  alldates = Array.from( new Set(data.map( d => d['boot_times']) ));
  alltime = Array.from( new Set(data.map( d => d['time_1']) ));
  alldates = alldates.sort( (a,b) => {
      return a - b;
  } );

  /* 按照开机次数分类 */
  sequantial = {}; 
  alldates.forEach( key => {
    sequantial[key] = []
  } );
  data.forEach( d => {
    sequantial[d['boot_times']].push(d)
  } );
  alldates.forEach( key => sequantial[key].sort(function(b,a){
      return b['top_times'] - a['top_times'];  //每个分组按照top大小排列
}));


render_init(data);

  let intervalId = setInterval(() => {
    let key = alldates[c];
    if(ChangeDay != 0){
        updataXAxis(sequantial[key]);
        ChangeDay = 0;  //把改变时间段的标志清零
    }
    if(click == 1){
        
        if(go != null){
          key = go;
          c = go - 1;
          updataXAxis(sequantial[key]);
        }
        
        render_update_alter(sequantial[key]);
        go = null;
        click = 0;
        aduration = 0;
    }
    
  }, aduration);
} );