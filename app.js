let trip={}

let dates=[]

let current=0

async function loadExcel(){

let res=await fetch("trip.xlsx")

let data=await res.arrayBuffer()

let wb=XLSX.read(data)

let sheet=wb.Sheets[wb.SheetNames[0]]

let json=XLSX.utils.sheet_to_json(sheet)

json.forEach(r=>{

let d=new Date(r.Date)

let key=d.toISOString().split("T")[0]

trip[key]={

location:r.Location,

hotel:r.Hotel,

spots:r.Spots? r.Spots.split(","):[],

remark:r.Remark||""

}

dates.push(key)

})

render()

}

function render(){

let date=dates[current]

let info=trip[date]

document.getElementById("dateTitle").innerText=date

document.getElementById("dayCount").innerText="Day "+(current+1)+" / "+dates.length

document.getElementById("location").innerText=info.location

document.getElementById("hotel").innerText=info.hotel

document.getElementById("remark").innerText=info.remark

let list=document.getElementById("spots")

list.innerHTML=""

info.spots.forEach(p=>{

let div=document.createElement("div")

div.className="item"

div.innerHTML=`

<span>${p}</span>

<button class="mapBtn" onclick="map('${p}')">Map</button>

`

list.appendChild(div)

})

}

function map(p){

window.open("https://www.google.com/maps/search/"+encodeURIComponent(p))

}

function prev(){

if(current>0){

current--

render()

}

}

function next(){

if(current<dates.length-1){

current++

render()

}

}

loadExcel()