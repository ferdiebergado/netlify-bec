"use strict";(()=>{var s=document.forms.namedItem("excelForm"),l=document.getElementById("excelFile");s&&l?s.addEventListener("submit",async a=>{a.preventDefault();let t=new FormData,n=l.files?.[0];if(n){t.append("excelFile",n);try{let e=await fetch("/convert",{method:"POST",body:t});if(e.ok){let i=await e.arrayBuffer(),d=new Uint8Array(i),m=new Blob([d],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),r=e.headers.get("Content-Disposition"),c=r&&r.match(/filename="(.+?)"/),f=c?c[1]:`em-${new Date().getTime()}.xlsx`,o=document.createElement("a");o.href=URL.createObjectURL(m),o.download=f,document.body.appendChild(o),o.click(),document.body.removeChild(o)}else console.error("Failed to convert. Server returned:",e.status,e.statusText)}catch(e){console.error("An error occurred during conversion:",e)}}else console.error("No file selected for conversion.")}):console.error("Form or file input not found.");})();
//# sourceMappingURL=app.js.map
