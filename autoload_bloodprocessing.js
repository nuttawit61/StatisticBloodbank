/* =====================================================================
 *  autoload_bloodprocessing.js — ห้อง Blood Processing
 *  สร้างโครงสร้าง BLOOD_LAB_DATA[ปี] จาก 5 ชีต:
 *    ขอเลือด -> RequestedBlood, เลือดหมดอายุ -> ExpiredBlood,
 *    Bl.processing -> BlProcessing(aboRh+antibodyId), FFP หมดอายุ -> FFPExpired,
 *    Antigentyping -> Antigentyping
 *  อัพเดทเฉพาะปีปัจจุบัน (2569) ; total ของ aboRh/antibody คำนวณจากผลรวม 12 เดือน
 * ===================================================================== */
(function () {
  "use strict";
  var CURRENT_YEAR = 2569;

  var MONTH = {}; // ชื่อเดือน -> index 0..11
  ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."].forEach(function(m,i){MONTH[m]=i;});
  ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"].forEach(function(m,i){MONTH[m]=i;});
  ["january","february","march","april","may","june","july","august","september","october","november","december"].forEach(function(m,i){MONTH[m]=i;});

  function n(v){ if(v===null||v===undefined||v==="")return 0; if(typeof v==="number")return v; var m=String(v).match(/-?\d+(?:\.\d+)?/); return m?Number(m[0]):0; }
  function monthIdxOf(v){ if(v==null)return -1; var k=String(v).trim().toLowerCase(); return (k in MONTH)?MONTH[k]:-1; }

  function parse(wb, H) {
    var XLSX=window.XLSX, G=function(s){return wb.Sheets[s]?H.sheetGrid(wb,s,XLSX):[];};
    var out={};

    // ---------- ขอเลือด -> RequestedBlood ----------
    (function(){
      var g=G("ขอเลือด"); if(!g.length)return;
      var arr=[]; for(var i=0;i<12;i++)arr.push({monthIdx:i+1,
        redCross:{O:0,A:0,B:0,AB:0,total:0}, others:{O:0,A:0,B:0,AB:0,total:0}, totalRBC:{O:0,A:0,B:0,AB:0,total:0}});
      g.forEach(function(row){
        var mi=monthIdxOf(row[0]); if(mi<0)return;
        arr[mi].redCross={O:n(row[1]),A:n(row[2]),B:n(row[3]),AB:n(row[4]),total:n(row[5])};
        arr[mi].others  ={O:n(row[6]),A:n(row[7]),B:n(row[8]),AB:n(row[9]),total:n(row[10])};
        arr[mi].totalRBC={O:n(row[11]),A:n(row[12]),B:n(row[13]),AB:n(row[14]),total:n(row[15])};
      });
      out.RequestedBlood=arr;
    })();

    // ---------- เลือดหมดอายุ -> ExpiredBlood ----------
    (function(){
      var g=G("เลือดหมดอายุ"); if(!g.length)return;
      var arr=[]; for(var i=0;i<12;i++)arr.push({monthIdx:i+1,
        expired:{O:0,A:0,B:0,AB:0,total:0}, components:{PRC:0,Adsol:0,LPB:0,PFB:0},
        donorOther:{O:0,A:0,B:0,AB:0,total:0}, rhNeg:""});
      g.forEach(function(row){
        var mi=monthIdxOf(row[1]); if(mi<0)return;   // เดือนอยู่คอลัมน์ B (index1)
        arr[mi].donorOther={O:n(row[2]),A:n(row[3]),B:n(row[4]),AB:n(row[5]),total:n(row[6])};
        arr[mi].components={PRC:n(row[10]),LPB:n(row[11]),PFB:n(row[12]),Adsol:n(row[13])};
        arr[mi].expired={O:n(row[14]),A:n(row[15]),B:n(row[16]),AB:n(row[17]),total:n(row[18])};
        arr[mi].rhNeg=(row[19]!=null&&row[19]!=="")?String(row[19]):"";
      });
      out.ExpiredBlood=arr;
    })();

    // ---------- FFP หมดอายุ -> FFPExpired ----------
    (function(){
      var g=G("FFP หมดอายุ"); if(!g.length)return;
      var arr=[]; for(var i=0;i<12;i++)arr.push({monthIdx:i+1,returnedWard:0,abnormalFFP:0,normalFFP:0,groups:{O:0,A:0,B:0,AB:0}});
      g.forEach(function(row){
        var mi=monthIdxOf(row[0]); if(mi<0)return;
        arr[mi].groups={A:n(row[2]),B:n(row[3]),AB:n(row[4]),O:n(row[5])};
        arr[mi].normalFFP=n(row[6]);
        arr[mi].abnormalFFP=n(row[7])+n(row[8])+n(row[9]);
        arr[mi].returnedWard=n(row[10]);
      });
      out.FFPExpired=arr;
    })();

    // ---------- Antigentyping ----------
    (function(){
      var g=G("Antigentyping"); if(!g.length)return;
      var arr=[]; for(var i=0;i<12;i++)arr.push({monthIdx:i+1,antigenMia:0,rhPhenotyping:0,g6pd:0,dat:0,datPos:0,datOther:0,cause:""});
      g.forEach(function(row){
        var mi=monthIdxOf(row[0]); if(mi<0)return;
        arr[mi].rhPhenotyping=n(row[1]); arr[mi].antigenMia=n(row[2]); arr[mi].g6pd=n(row[3]);
        arr[mi].dat=n(row[4]); arr[mi].datPos=n(row[5]); arr[mi].datOther=n(row[6]);
        arr[mi].cause=(row[7]!=null&&row[7]!=="")?String(row[7]).trim():"";
      });
      out.Antigentyping=arr;
    })();

    // ---------- Bl.processing -> BlProcessing ----------
    (function(){
      var g=G("Bl.processing"); if(!g.length)return;
      var norm=H.norm;
      // aboRh: ตารางแรก group(col0)+rh(col1)+12 เดือน(col2-13) จนถึง "Total"
      var aboRh=[], curGroup="";
      var startA=-1;
      for(var i=0;i<g.length;i++){ var b=norm(g[i][1]); if(b==="abo,rh"||(norm(g[i][0])==="bloodgroup")){startA=i+1;break;} }
      if(startA>=0){
        for(var r=startA;r<g.length;r++){
          var c0=(g[r][0]!=null)?String(g[r][0]).trim():"";
          var rh=(g[r][1]!=null)?String(g[r][1]).trim():"";
          if(norm(c0)==="total"||norm(c0)==="test") break;
          if(c0!=="") curGroup=c0.replace(/\s+/g,"");
          if(rh===""){ if(c0==="")continue; else continue; }
          var months=[]; var sum=0;
          for(var m=0;m<12;m++){ var v=n(g[r][2+m]); months.push(v); sum+=v; }
          aboRh.push({group:curGroup,rh:rh,months:months,total:sum});
        }
      }
      // antibodyId: หัวข้อ "Antibody" + 12 เดือน(col1-12) ; [0]=header
      var antibodyId=[];
      var hdr=-1;
      for(var j=0;j<g.length;j++){ if(norm(g[j][0])==="antibody"){hdr=j;break;} }
      if(hdr>=0){
        antibodyId.push({name:"Antibody",months:["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."],total:"รวม"});
        for(var k=hdr+1;k<g.length;k++){
          var nm=(g[k][0]!=null)?String(g[k][0]).trim():"";
          if(nm===""||norm(nm)==="total"||norm(nm).indexOf("รวม")>=0) break;
          var mo=[],s2=0; for(var mm=0;mm<12;mm++){var vv=n(g[k][1+mm]);mo.push(vv);s2+=vv;}
          antibodyId.push({name:nm,months:mo,total:s2});
        }
      }
      out.BlProcessing={antibodyId:antibodyId,aboRh:aboRh};
    })();

    return out;
  }

  function matchYear(filename){ return /^สถิติ\s*69\.xlsx$/i.test(String(filename))?CURRENT_YEAR:null; }
  function applyYear(year,data){
    if(year!==CURRENT_YEAR||typeof BLOOD_LAB_DATA==="undefined"||!BLOOD_LAB_DATA[String(year)])return;
    var obj=BLOOD_LAB_DATA[String(year)];
    Object.keys(data).forEach(function(k){ obj[k]=data[k]; });
  }
  function render(){ if(typeof refresh==="function")refresh(); if(typeof buildCharts==="function")buildCharts(); }

  DashboardAutoload.register({
    mode:"directory", storeKey:"bloodprocessing", snapshotFile:"live_bloodprocessing.js",
    defaultYear:CURRENT_YEAR, match:matchYear, parse:parse, applyYear:applyYear, render:render
  });
})();
