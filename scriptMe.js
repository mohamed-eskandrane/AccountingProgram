const sheetId = '1FQ3-PozLrgcj1vwHiGHS0ToIVetuhWonrUMANE8-N94';
const base = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?`;
let query = encodeURIComponent('Select *');
let Users="Users";
let UrlUsers = `${base}&sheet=${Users}&tq=${query}`;
let DataUsers = [];
let Accounts="Accounts";
let UrlAccounts = `${base}&sheet=${Accounts}&tq=${query}`;
let DataAccounts = [];
let WareHouse="WareHouse";
let UrlWareHouse = `${base}&sheet=${WareHouse}&tq=${query}`;
let DataWareHouse = [];
let Mats="Mats";
let UrlMat = `${base}&sheet=${Mats}&tq=${query}`;
let DataMat = [];
let Purchases="Purchases";
let UrlPurchases = `${base}&sheet=${Purchases}&tq=${query}`;
let DataPurchases = [];
let Move="Move";
let UrlMove = `${base}&sheet=${Move}&tq=${query}`;
let DataMove = [];
let Sales="Sales";
let UrlSales = `${base}&sheet=${Sales}&tq=${query}`;
let DataSales = [];
let Entry="Entry";
let UrlEntry = `${base}&sheet=${Entry}&tq=${query}`;
let DataEntry = [];
let Teacher="Teacher";
let UrlTeacher = `${base}&sheet=${Teacher}&tq=${query}`;
let DataTeacher = [];
let AccountsSeting="AccountsSeting";
let UrlAccountsSeting = `${base}&sheet=${AccountsSeting}&tq=${query}`;
let DataAccountsSeting = [];
document.addEventListener('DOMContentLoaded', init)
function init() {
  ConvertMode();
  LoadUsers();
  if (typeof(Storage) !== "undefined") {
    if( localStorage.getItem("PassWord")!=null){
      document.getElementById("User_PassWord").value=localStorage.getItem("PassWord");
    }
    if( localStorage.getItem("User_Index")!=null){
      ShowSelectForm(localStorage.getItem("ActiveForm"));
      document.getElementById("Myusername").value=localStorage.getItem("User_Name");
      LoadAccounts();
      LoadAccountsSeting();
      LoadWareHouse();
      LoadMat();
      LoadPurchases();
      LoadMove();
      LoadSales();
      LoadEntry();
      LoadTeacher();
    }
  }
}
function ShowSelectForm(ActiveForm){
  document.getElementById("loginPage").style.display="none";
  document.getElementById("Main").style.display="none";
  document.getElementById("PurchasesWi").style.display="none";
  document.getElementById("PurchasesBrowser").style.display="none";
  document.getElementById("SalesWi").style.display="none";
  document.getElementById("SalesBrowser").style.display="none";
  document.getElementById("EntryWi").style.display="none";
  document.getElementById("EntryBrowser").style.display="none";
  document.getElementById("AccountBrowser").style.display="none";
  document.getElementById("MatsBrowser").style.display="none";
  document.getElementById(ActiveForm).style.display="flex";
  localStorage.setItem("ActiveForm",ActiveForm)
}
function GetDateFromString(Str){
  let MM,DD;
  let ZZ=[];
  let SS=String(Str).substring(5,String(Str).length-1);
  ZZ=SS.split(",");
  if (Number(ZZ[1])<9 && Number(ZZ[1]).length!= 2){ MM=0 + String(parseInt(ZZ[1]) + 1)}else{ MM=(parseInt(ZZ[1]) + 1)}
  if (Number(ZZ[2])<=9 && Number(ZZ[1]).length!= 2){ DD=0 + ZZ[2]}else{ DD=ZZ[2]}
  return ZZ[0] + "-" + MM + "-" + DD
}
// *************************************Main**************
function ShowPurchasesBrowser(){
  let Index=localStorage.getItem("User_Index");
  if (DataUsers[Index].IsAdmin=="yes"){
    ShowSelectForm("PurchasesBrowser");
    LoadPurchasesToTable();
  }
}

function ShowSalesBrowser(){
  let Index=localStorage.getItem("User_Index");
  if (DataUsers[Index].IsAdmin=="yes"){
    ShowSelectForm("SalesBrowser");
    LoadSalesToTable();
  }
}

function ShowEntryBrowser(){
  let Index=localStorage.getItem("User_Index");
  if (DataUsers[Index].IsAdmin=="yes"){
    ShowSelectForm("EntryBrowser");
    LoadEntryToTable();
  }
}

function ShowAccountsBrowser(){
  let Index=localStorage.getItem("User_Index");
  if (DataUsers[Index].IsAdmin=="yes"){
  ShowSelectForm("AccountBrowser");
  LoadAccountsToTable();
  }
}

function ShowMatsBrowser(){
  let Index=localStorage.getItem("User_Index");
  if (DataUsers[Index].IsAdmin=="yes"){
  ShowSelectForm("MatsBrowser");
  LoadMatsToTable();
  }
}

function SignOutUser(){
  localStorage.removeItem("User_Index");
  localStorage.removeItem("User_Name");
  document.getElementById('Myusername').value="";
  ShowSelectForm("loginPage");
}
function GoToMain(){
  ShowSelectForm("Main");
}
// **********************SignIN*****************
function LoadUsers(){
  DataUsers=[];
  fetch(UrlUsers)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataUsers.push(rowUser);
      })
  })
}
function LoadAccounts(){
  DataAccounts=[];
  fetch(UrlAccounts)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataAccounts.push(rowUser);
      })
      LoadAccountName()
  })
}
function LoadAccountName(){
  let AccountName,AccountNum;
  let optionClass;
  let listAccountName =document.getElementById("listAccountName");
  let listAccountNameS =document.getElementById("listAccountNameS");
  listAccountName.innerHTML="";
  for (let index = 0; index < DataAccounts.length; index++) {
    AccountNum=DataAccounts[index].AccountNum
    AccountName=DataAccounts[index].AccountName
    if(AccountNum!=""){
      optionClass=document.createElement("option");
      optionClass.value=AccountName;
      optionClass.textContent=AccountName;
      listAccountName.appendChild(optionClass);
      optionClass=document.createElement("option");
      optionClass.value=AccountName;
      optionClass.textContent=AccountName;
      listAccountNameS.appendChild(optionClass);
    }
  }
}
function LoadWareHouse(){
  DataWareHouse=[];
  fetch(UrlWareHouse)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataWareHouse.push(rowUser);
      })
  })
}
function LoadWareHouseName(Num){
  let WareHouseName,NumberH;
  let optionClass;
  let listMatName =document.getElementById("PRBWareHP" + Num);
  listMatName.innerHTML="";
  for (let index = 0; index < DataWareHouse.length; index++) {
    NumberH=DataWareHouse[index].Num
    WareHouseName=DataWareHouse[index].WareHouseName
    if(NumberH!=""){
      optionClass=document.createElement("option");
      optionClass.value=WareHouseName;
      optionClass.textContent=WareHouseName;
      listMatName.appendChild(optionClass);
    }
  }
}
function LoadWareHouseNameS(Num){
  let WareHouseName,NumberH;
  let optionClass;
  let listMatName =document.getElementById("PRBWareHS" + Num);
  listMatName.innerHTML="";
  for (let index = 0; index < DataWareHouse.length; index++) {
    NumberH=DataWareHouse[index].Num
    WareHouseName=DataWareHouse[index].WareHouseName
    if(NumberH!=""){
      optionClass=document.createElement("option");
      optionClass.value=WareHouseName;
      optionClass.textContent=WareHouseName;
      listMatName.appendChild(optionClass);
    }
  }
}
function LoadMat(){
  DataMat=[];
  fetch(UrlMat)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataMat.push(rowUser);
      })
  })
}
function LoadMove(){
  DataMove=[];
  fetch(UrlMove)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataMove.push(rowUser);
      })
  })
}
function LoadMatsName(Num){
  let MatName,MatCode;
  let optionClass;
  let listMatName =document.getElementById("listMatName" + Num);
  listMatName.innerHTML="";
  for (let index = 0; index < DataMat.length; index++) {
    MatCode=DataMat[index].MatCode
    MatName=DataMat[index].MatName
    if(MatCode!=""){
      optionClass=document.createElement("option");
      optionClass.value=MatName;
      optionClass.textContent=MatName;
      listMatName.appendChild(optionClass);
    }
  }
}
function IsfoundUser(TPassWord){
  let error_User_ID= document.getElementById("error_User_ID");
    for (let index = 0; index < DataUsers.length; index++) {
      if(TPassWord==DataUsers[index].PassWord){
        localStorage.setItem("User_Index", index);
        return true;
      }
    }
      error_User_ID.className="fa fa-warning";
      return false ;
  }

  function foundIndex(TPassWord){
      for (let index = 0; index < DataUsers.length; index++) {
        if(TPassWord==DataUsers[index].PassWord){
          return index
        }
      }
      return -1
    }
  
function Istrue(TPassWord){
  let error_User_ID= document.getElementById("error_User_ID");
  if(TPassWord===""){ error_User_ID.className="fa fa-warning"; return false;}else{ error_User_ID.className="" }
  if(IsfoundUser(TPassWord)===false){return false}else{error_User_ID.className=""}
  return true;
}

function Sign_In(){
  let User_PassWord= document.getElementById("User_PassWord");
  if (Istrue(User_PassWord.value)===true){
    let User_Index = localStorage.getItem("User_Index");
    localStorage.setItem("User_Name", DataUsers[User_Index].UserName);
    localStorage.setItem("PassWord",DataUsers[User_Index].PassWord);
    document.getElementById("Myusername").value=localStorage.getItem("User_Name");
    ShowSelectForm("Main");
  }
}

function ShowPassword(){
  let User_PassWord= document.getElementById("User_PassWord");
  let Eye_Password= document.getElementById("Eye_Password");
  if (Eye_Password.className=="fa fa-eye"){
    User_PassWord.type="text";
    Eye_Password.className="fa fa-eye-slash";
  }else{
    User_PassWord.type="password";
    Eye_Password.className="fa fa-eye";
  }
}


// ********************PurchasesWi
function LoadPurchases(){
  DataPurchases=[];
  fetch(UrlPurchases)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataPurchases.push(rowUser);
      })
  })
}
function GetNameAccountFormCode(Code){
  for (let index = 0; index < DataAccounts.length; index++) {
    if(DataAccounts[index].AccountCode==Code ){
      document.getElementById("AccountNameP").value=DataAccounts[index].AccountName;
      return ;
    }
  }
  document.getElementById("AccountNameP").value="";
}

function GetCodeAccountFormName(Nam){
  for (let index = 0; index < DataAccounts.length; index++) {
    if(DataAccounts[index].AccountName==Nam ){
      document.getElementById("AccountCodeP").value=DataAccounts[index].AccountCode;
      return ;
    }
  }
  document.getElementById("AccountCodeP").value="";
}

function GetCodeMatFormName(NameA, TRow){
  let Row=String(TRow).slice(8,String(TRow).length)
  for (let index = 0; index < DataMat.length; index++) {
    if(DataMat[index].MatName==NameA ){
      document.getElementById("PRBCodeP" + Row).value=DataMat[index].MatCode;
      return ;
    }
  }
  document.getElementById("PRBCodeP" + Row).value="";
}

function GetNameMatFormCode(Code , TRow){
  let Row=String(TRow).slice(8,String(TRow).length);
  for (let index = 0; index < DataMat.length; index++) {
    if(DataMat[index].MatCode==Code ){
      document.getElementById("PRBCodeN" + Row).value=DataMat[index].MatName;
      return ;
    }
  }
  document.getElementById("PRBCodeN" + Row).value="";
}

function CalculateRowA(Am,TRow){
  let Row=String(TRow).slice(10,String(TRow).length);
  let Pr=document.getElementById("PRBPriceP" + Row).value
  document.getElementById("PRBMatTotalP" + Row).value = Pr * Am
  CaluclateTotal();
}
function CalculateRowP(Pr,TRow){
  let Row=String(TRow).slice(9,String(TRow).length);
  let Am=document.getElementById("PRBAmountP" + Row).value
  document.getElementById("PRBMatTotalP" + Row).value = Pr * Am
  CaluclateTotal();
}

function CaluclateTotal(){
  let PTotal = document.getElementById("PTotal");
  let PTotalNet = document.getElementById("PTotalNet");
  let ShipPr = document.getElementById("ShipPr");
  let CountRow= document.getElementById("CountRow1");
  let CountMat= document.getElementById("CountMat");
  let TableInputMt=document.getElementsByClassName("TableInputMt");
  let total = 0 ;
  let Rows = 0 ;
  let totalMats = 0 ;
  for (let index = 0; index < TableInputMt.length; index++) {
    if (TableInputMt.item(index).parentElement.parentElement.hidden==false){
      total =  total + Number(TableInputMt.item(index).value);
      Rows= Rows + 1
    }
  }
  PTotal.value=total;
  CountRow.value=Rows;
  PTotalNet.value= total + Number(ShipPr.value);
  let TableInputMa=document.getElementsByClassName("TableInputMa");
  for (let index = 0; index < TableInputMa.length; index++) {
    if (TableInputMa.item(index).parentElement.parentElement.hidden==false){
      totalMats =  totalMats + Number(TableInputMa.item(index).value);
    }
  }
  CountMat.value= totalMats;
}

function AddRowPrElement() {
  var td;
  let bodydata=document.getElementById("bodyBillPur");
  let row = bodydata.insertRow();
  row.id="PRB" + bodydata.childElementCount;
  row.className="PRBRow";
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "text";
  btb.autocomplete="off";
  btb.className="TableInputMc";
  btb.style.width="89%";
  btb.id="PRBCodeP" + bodydata.childElementCount ;
  btb.name="MatCode" + bodydata.childElementCount ;
  btb.onkeyup=function(){GetNameMatFormCode(this.value,this.id)};
  td.appendChild(btb)
  row.appendChild(td=document.createElement('td'));
  td.innerHTML=`<input list='listMatName${bodydata.childElementCount}' id="PRBCodeN${bodydata.childElementCount}"  style='width:95%;' class="TableInputMn" onchange="GetCodeMatFormName(this.value,this.id)" autocomplete="off">`;
  var datalist = document.createElement('datalist');
  datalist.id=`listMatName${bodydata.childElementCount}`;
  td.appendChild(datalist);
  LoadMatsName(bodydata.childElementCount);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "number";
  btb.autocomplete="off";
  btb.className="TableInputMa";
  btb.style.width="92%";
  btb.id="PRBAmountP" + bodydata.childElementCount ;
  btb.name="Amount" + bodydata.childElementCount ;
  btb.onkeyup=function(){CalculateRowA(this.value,this.id)};
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('select');
  btb.className="TableInputMw";
  btb.style.width="92%";
  btb.id="PRBWareHP" + bodydata.childElementCount ;
  btb.name="WareHouse" + bodydata.childElementCount ;
  td.appendChild(btb);
  LoadWareHouseName(bodydata.childElementCount);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "number";
  btb.autocomplete="off";
  btb.className="TableInputMp";
  btb.style.width="92%";
  btb.id="PRBPriceP" + bodydata.childElementCount ;
  btb.name="Price" + bodydata.childElementCount ;
  btb.onkeyup=function(){CalculateRowP(this.value,this.id)};
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "number";
  btb.autocomplete="off";
  btb.className="TableInputMt";
  btb.style.width="92%";
  btb.readOnly=true;
  btb.id="PRBMatTotalP" + bodydata.childElementCount ;
  btb.name="MatTotal" + bodydata.childElementCount ;
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "text";
  btb.autocomplete="off";
  btb.className="TableInputMno";
  btb.style.width="92%";
  btb.id="PRBNoteP" + bodydata.childElementCount ;
  btb.name="Note" + bodydata.childElementCount ;
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  btb = document.createElement('input');
  btb.type = "button";
  btb.id="PRBD" + bodydata.childElementCount;
  btb.value = "X";
  btb.className="BtnStyle";
  btb.onclick=function(){DeleteRow(this.id)};
  td.appendChild(btb)
  };

  function DeleteRow(TRow){
    let Row=String(TRow).slice(4,String(TRow).length);
    document.getElementById("PRB" + Row).hidden=true;
    CaluclateTotal();
  }


function IstrueDataInform(){
  let BillDateP=document.getElementById("BillDateP");
  let AccountCodeP=document.getElementById("AccountCodeP");
  let AccountNameP=document.getElementById("AccountNameP");
  let bodyBillPur=document.getElementById("bodyBillPur"); 
  let PurchasesItems=document.getElementById("PurchasesItems");   
  if(BillDateP.value==""){BillDateP.style.border="2px solid #ff0000";return false}else{BillDateP.style.border="none";}
  if(AccountCodeP.value==""){AccountCodeP.style.border="2px solid #ff0000";return false}else{AccountCodeP.style.border="none";}
  if(AccountNameP.value==""){AccountNameP.style.border="2px solid #ff0000";return false}else{AccountNameP.style.border="none";}
  if(bodyBillPur.childElementCount==0){PurchasesItems.style.border="2px solid #ff0000";return false}else{PurchasesItems.style.border="2px solid rgb(155, 153, 153)";}
  if(IsTableInputMcTrue()==false){return false};
  if(IsTableInputMnTrue()==false){return false};
  if(IsTableInputMaTrue()==false){return false};
  if(IsTableInputMpTrue()==false){return false};
  RemoveRowHidden();
  if(bodyBillPur.childElementCount==0){PurchasesItems.style.border="2px solid #ff0000";return false}else{PurchasesItems.style.border="2px solid rgb(155, 153, 153)";}
  return true;
}

function RemoveRowHidden(){
  let ArraryNamesRow=[];
  let ArraryNames=[];
    let PRBRow=document.getElementsByClassName("PRBRow");
    for (let index1 = PRBRow.length - 1; index1 >=0; index1--) {
      if (PRBRow.item(index1).hidden==true){
        PRBRow.item(index1).remove();
      }
    }
    for (let index = 0; index < PRBRow.length; index++) {
        var RowN=PRBRow.item(index).children;
        for (let indexN = 0; indexN < RowN.length; indexN++) {
          var RowNi =RowN.item(indexN).children.item(0)
          if(RowNi.name !=""){
            ArraryNames.push(RowNi.name);
          }
        }
        ArraryNamesRow.push(ArraryNames);
        ArraryNames=[];
    }
    document.getElementById("ArrayTableP").value=ArraryNamesRow;
}

function IsTableInputMpTrue(){
  let TableInputMp=document.getElementsByClassName("TableInputMp");
  for (let index = 0; index < TableInputMp.length; index++) {
    if (TableInputMp.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMp.item(index).value==""){
        TableInputMp.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMp.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}

function IsTableInputMaTrue(){
  let TableInputMa=document.getElementsByClassName("TableInputMa");
  for (let index = 0; index < TableInputMa.length; index++) {
    if (TableInputMa.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMa.item(index).value==""){
        TableInputMa.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMa.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}

function IsTableInputMcTrue(){
  let TableInputMc=document.getElementsByClassName("TableInputMc");
  for (let index = 0; index < TableInputMc.length; index++) {
    if (TableInputMc.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMc.item(index).value==""){
        TableInputMc.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMc.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}
function IsTableInputMnTrue(){
  let TableInputMn=document.getElementsByClassName("TableInputMn");
  for (let index = 0; index < TableInputMn.length; index++) {
    if (TableInputMn.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMn.item(index).value==""){
        TableInputMn.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMn.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}
function onsubmitForm1(){
  if(IstrueDataInform()===true){
    document.getElementById("BillNumberP").value=MaxBillNumber();
    document.getElementById("ModeP").value="1";
    onsubmitForm(5000);
  }
}

function BillNumberIsfound(BillNumber){
    for (let index = 0; index < DataPurchases.length; index++) {
      if(DataPurchases[index].BillNumber==BillNumber ){
        return true;
      }
  }
  return false;
}

function onsubmitForm2(){
  let BillNumberP =document.getElementById("BillNumberP")
  if(IstrueDataInform()===true){
    if(BillNumberIsfound(BillNumberP.value)==true){
      document.getElementById("ModeP").value="2";
      onsubmitForm(10000);
    }
  }
}

function onsubmitForm3(){
  let BillNumberP =document.getElementById("BillNumberP")
  if(BillNumberIsfound(BillNumberP.value)==true){
  document.getElementById("ModeP").value="3";
  onsubmitForm(8000);
  }
}

function onsubmitForm(Time){
  document.getElementById("typeP").value="1";
  document.getElementById("UserNameP").value=localStorage.getItem("User_Name");
  let MainForm=document.getElementById("FormPurchasesDetails");
  var w = window.open('', 'form_target', 'width=600, height=400');
  MainForm.target = 'form_target';
  MainForm.action='https://script.google.com/macros/s/AKfycbzrvVLAeb6YwTxzwZq6nSLuerWq4qnUFuY2_7o2NO76gqHAlna6NC9wmBEXBFiSrzjScw/exec'
  MainForm.submit();
  if (MainForm.onsubmit()==true){
    const myTimeout = setTimeout(function(){ 
                w.close();
                clearTimeout(myTimeout)
                location.reload();
    }, Time);
  }
} 

// **************************PurchasesBrowser***********
function GetNameAccount(Code){
  for (let index = 0; index < DataAccounts.length; index++) {
    if(DataAccounts[index].AccountCode==Code ){
      return DataAccounts[index].AccountName
    }
  }
  return "none"
}

function LoadPurchasesToTable(){
  const myTimeout = setTimeout(function(){ 
  document.getElementById("bodydata").innerHTML=""
  if (isNaN(DataPurchases[0].Num)){return}
  for (let index = 0; index < DataPurchases.length; index++) {
    if(DataPurchases[index].Num!="" ){
      AddRowPr(DataPurchases[index].BillNumber,DataPurchases[index].EDate,DataPurchases[index].AccountCode,DataPurchases[index].Total,DataPurchases[index].Note,DataPurchases[index].UserName)
    }
  }
  clearTimeout(myTimeout)
}, 1000);
}

function AddRowPr(BillNumberP,BillDateP,AccountCodeP,TotalP,NoteP,UserNameP) {
  let bodydata=document.getElementById("bodydata");
  let row = bodydata.insertRow();
  row.id="PR" + bodydata.childElementCount;
  let cell = row.insertCell();
  cell.id="PR" + bodydata.childElementCount + "BillNumberP";
  cell.innerHTML = BillNumberP;
  cell = row.insertCell();
  cell.id="PR" + bodydata.childElementCount + "BillDateP";
  cell.innerHTML = GetDateFromString(BillDateP);
  cell = row.insertCell();
  cell.id="PR" + bodydata.childElementCount + "AccountCodeP";
  cell.innerHTML = AccountCodeP;
  cell = row.insertCell();
  cell.id="PR" + bodydata.childElementCount + "AccountNameP";
  cell.innerHTML = GetNameAccount(AccountCodeP);
  cell = row.insertCell();
  cell.id="PR" + bodydata.childElementCount + "TotalP";
  cell.innerHTML = TotalP;
  cell = row.insertCell();
  cell.id="PR" + bodydata.childElementCount + "NoteP";
  cell.innerHTML = NoteP;
  cell = row.insertCell();
  cell.id="PR" + bodydata.childElementCount + "UserNameP";
  cell.innerHTML = UserNameP;
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "button";
  btb.id="ButPr" + bodydata.childElementCount;
  btb.value = "تعديل";
  td.appendChild(btb)
  btb.style.cursor="pointer";
  btb.style.color="red";
  btb.style.padding="2px";
  btb.onclick=function(){showdatarows()};
  };

function MaxBillNumber(){
  let XX;
  let BillNumbers=[];
  for (let index = 0; index < DataPurchases.length; index++) {
    if(DataPurchases[index].Num!="" ){
      BillNumbers.push(DataPurchases[index].BillNumber);
    }
  }
  XX= Math.max.apply(Math, BillNumbers) + 1;
  if(isNaN(XX)==true){return 1}else{return XX}
}
  function ShowDetails(){
    ShowSelectForm("PurchasesWi");
    document.getElementById("BillNumberP").value=MaxBillNumber();
    document.getElementById("BillDateP").value=""
    document.getElementById("AccountCodeP").value=""
    document.getElementById("AccountNameP").value=""
    document.getElementById("NoteP").value=""
    document.getElementById("PTotal").value=""
    document.getElementById("PTotalNet").value=""
    document.getElementById("bodyBillPur").innerHTML=""    
  }
  function GetMatNameFormCode1(Code){
    for (let index = 0; index < DataMat.length; index++) {
      if(DataMat[index].MatCode==Code ){
        return DataMat[index].MatName;
      }
    }
  }
  function showdatarows() {
    let indextable= document.activeElement.parentElement.parentElement.id;
    let IndexRow =document.getElementById(indextable).children.item(0).textContent  ;
    let NameAccountTable=document.getElementById(indextable).children.item(3).textContent  ;
    ShowDetails();
    for (let index = 0; index < DataPurchases.length; index++) {
      if(DataPurchases[index].BillNumber == IndexRow){
        document.getElementById("BillNumberP").value= IndexRow;
        document.getElementById("BillDateP").value=GetDateFromString(DataPurchases[index].EDate);
        document.getElementById("AccountCodeP").value=DataPurchases[index].AccountCode;
        document.getElementById("AccountNameP").value=NameAccountTable;
        document.getElementById("NoteP").value=DataPurchases[index].Note;
        document.getElementById("ShipPr").value=DataPurchases[index].Ship;
        document.getElementById("PTotalNet").value=DataPurchases[index].Total;
      }
    }
    let IndexTableRow = 1;
    for (let indexM = 0; indexM < DataMove.length; indexM++) {
      if(DataMove[indexM].BillNumber == IndexRow  && DataMove[indexM].Type=="1"){
        AddRowPrElement();
        document.getElementById(`PRBCodeP${IndexTableRow}`).value= DataMove[indexM].MatCode;
        document.getElementById(`PRBCodeN${IndexTableRow}`).value=GetMatNameFormCode1(DataMove[indexM].MatCode)
        document.getElementById(`PRBAmountP${IndexTableRow}`).value= DataMove[indexM].Amount;
        document.getElementById(`PRBWareHP${IndexTableRow}`).value= DataMove[indexM].WareHouse;
        document.getElementById(`PRBPriceP${IndexTableRow}`).value= DataMove[indexM].Price;
        document.getElementById(`PRBMatTotalP${IndexTableRow}`).value= DataMove[indexM].MatTotal;
        document.getElementById(`PRBNoteP${IndexTableRow}`).value= DataMove[indexM].Note;
        IndexTableRow++;
      }
    }
    CaluclateTotal();
    document.getElementById("CountRow").value=document.getElementById("CountRow1").value
  };

// ********************SalesWi
function LoadSales(){
  DataSales=[];
  fetch(UrlSales)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataSales.push(rowUser);
      })
  })
}
function GetNameAccountFormCodeS(Code){
  for (let index = 0; index < DataAccounts.length; index++) {
    if(DataAccounts[index].AccountCode==Code ){
      document.getElementById("AccountNameS").value=DataAccounts[index].AccountName;
      document.getElementById("Discount").value=DataAccounts[index].Discount;
      return ;
    }
  }
  document.getElementById("AccountNameS").value="";
  document.getElementById("Discount").value="";
}

function GetCodeAccountFormNameS(Nam){
  for (let index = 0; index < DataAccounts.length; index++) {
    if(DataAccounts[index].AccountName==Nam ){
      document.getElementById("AccountCodeS").value=DataAccounts[index].AccountCode;
      document.getElementById("Discount").value=DataAccounts[index].Discount;
      return ;
    }
  }
  document.getElementById("AccountCodeS").value="";
  document.getElementById("Discount").value="";
}

function GetCodeMatFormNameS(NameA, TRow){
  let Row=String(TRow).slice(9,String(TRow).length)
  for (let index = 0; index < DataMat.length; index++) {
    if(DataMat[index].MatName==NameA ){
      document.getElementById("PRBCodeS" + Row).value=DataMat[index].MatCode;
      return ;
    }
  }
  document.getElementById("PRBCodeS" + Row).value="";
}

function GetNameMatFormCodeS(Code , TRow){
  let Row=String(TRow).slice(8,String(TRow).length);
  for (let index = 0; index < DataMat.length; index++) {
    if(DataMat[index].MatCode==Code ){
      document.getElementById("PRBCodeNS" + Row).value=DataMat[index].MatName;
      return ;
    }
  }
  document.getElementById("PRBCodeNS" + Row).value="";
}

function CalculateRowAS(Am,TRow){
  let Row=String(TRow).slice(10,String(TRow).length);
  let Pr=document.getElementById("PRBPriceS" + Row).value
  document.getElementById("PRBMatTotalS" + Row).value = Pr * Am
  CaluclateTotalS();
}
function CalculateRowS(Pr,TRow){
  let Row=String(TRow).slice(9,String(TRow).length);
  let Am=document.getElementById("PRBAmountS" + Row).value
  document.getElementById("PRBMatTotalS" + Row).value = Pr * Am
  CaluclateTotalS();
}

function CaluclateTotalS(){
  let PTotal = document.getElementById("STotal");
  let ShipS = document.getElementById("ShipS");
  let Discount = document.getElementById("Discount");
  let DiscountPaid = document.getElementById("DiscountPaid");
  let STotalNet = document.getElementById("STotalNet");
  let CountRow= document.getElementById("CountRowS1");
  let CountMat= document.getElementById("CountMatS");
  let TableInputMt=document.getElementsByClassName("TableInputMtS");
  let total = 0 ;
  let Rows = 0 ;
  let totalMats = 0 ;
  for (let index = 0; index < TableInputMt.length; index++) {
    if (TableInputMt.item(index).parentElement.parentElement.hidden==false){
      total =  total + Number(TableInputMt.item(index).value);
      Rows= Rows + 1
    }
  }
  let Dis=0;
  if(Discount.value > 0.99){
    Dis=Discount.value;
  }else{
    Dis=Discount.value * total;
  }
  PTotal.value=total;
  CountRow.value=Rows;
  STotalNet.value=  total +  Number(ShipS.value) - DiscountPaid.value * (total - Dis ) - Dis ;
  let TableInputMa=document.getElementsByClassName("TableInputMaS");
  for (let index = 0; index < TableInputMa.length; index++) {
    if (TableInputMa.item(index).parentElement.parentElement.hidden==false){
      totalMats =  totalMats + Number(TableInputMa.item(index).value);
    }
  }
  CountMat.value= totalMats;
}

function AddRowPrElementS() {
  var td;
  let bodydata=document.getElementById("bodyBillSa");
  let row = bodydata.insertRow();
  row.id="PRBS" + bodydata.childElementCount;
  row.className="PRBRowS";
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "text";
  btb.autocomplete="off";
  btb.className="TableInputMcS";
  btb.style.width="89%";
  btb.id="PRBCodeS" + bodydata.childElementCount ;
  btb.name="MatCode" + bodydata.childElementCount ;
  btb.onkeyup=function(){GetNameMatFormCodeS(this.value,this.id)};
  td.appendChild(btb)
  row.appendChild(td=document.createElement('td'));
  td.innerHTML=`<input list='listMatName${bodydata.childElementCount}' id="PRBCodeNS${bodydata.childElementCount}"  style='width:95%;' class="TableInputMnS" onchange="GetCodeMatFormNameS(this.value,this.id)" autocomplete="off">`;
  var datalist = document.createElement('datalist');
  datalist.id=`listMatName${bodydata.childElementCount}`;
  td.appendChild(datalist);
  LoadMatsName(bodydata.childElementCount);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "number";
  btb.autocomplete="off";
  btb.className="TableInputMaS";
  btb.style.width="92%";
  btb.id="PRBAmountS" + bodydata.childElementCount ;
  btb.name="Amount" + bodydata.childElementCount ;
  btb.onkeyup=function(){CalculateRowAS(this.value,this.id)};
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('select');
  btb.className="TableInputMwS";
  btb.style.width="92%";
  btb.id="PRBWareHS" + bodydata.childElementCount ;
  btb.name="WareHouse" + bodydata.childElementCount ;
  td.appendChild(btb);
  LoadWareHouseNameS(bodydata.childElementCount);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "number";
  btb.autocomplete="off";
  btb.className="TableInputMpS";
  btb.style.width="92%";
  btb.id="PRBPriceS" + bodydata.childElementCount ;
  btb.name="Price" + bodydata.childElementCount ;
  btb.onkeyup=function(){CalculateRowS(this.value,this.id)};
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "number";
  btb.autocomplete="off";
  btb.className="TableInputMtS";
  btb.style.width="92%";
  btb.readOnly=true;
  btb.id="PRBMatTotalS" + bodydata.childElementCount ;
  btb.name="MatTotal" + bodydata.childElementCount ;
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "text";
  btb.autocomplete="off";
  btb.className="TableInputMnoS";
  btb.style.width="92%";
  btb.id="PRBNoteS" + bodydata.childElementCount ;
  btb.name="Note" + bodydata.childElementCount ;
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  btb = document.createElement('input');
  btb.type = "button";
  btb.id="PRBDS" + bodydata.childElementCount;
  btb.value = "X";
  btb.className="BtnStyle";
  btb.onclick=function(){DeleteRowS(this.id)};
  td.appendChild(btb)
  };


  function DeleteRowS(TRow){
    let Row=String(TRow).slice(5,String(TRow).length);
    document.getElementById("PRBS" + Row).hidden=true;
    CaluclateTotalS();
  }


  
function IstrueDataInformS(){
  let BillDateP=document.getElementById("BillDateS");
  let AccountCodeP=document.getElementById("AccountCodeS");
  let AccountNameP=document.getElementById("AccountNameS");
  let bodyBillPur=document.getElementById("bodyBillSa"); 
  let PurchasesItems=document.getElementById("SalesItems");   
  if(BillDateP.value==""){BillDateP.style.border="2px solid #ff0000";return false}else{BillDateP.style.border="none";}
  if(AccountCodeP.value==""){AccountCodeP.style.border="2px solid #ff0000";return false}else{AccountCodeP.style.border="none";}
  if(AccountNameP.value==""){AccountNameP.style.border="2px solid #ff0000";return false}else{AccountNameP.style.border="none";}
  if(bodyBillPur.childElementCount==0){PurchasesItems.style.border="2px solid #ff0000";return false}else{PurchasesItems.style.border="2px solid rgb(155, 153, 153)";}
  if(IsTableInputMcTrueS()==false){return false};
  if(IsTableInputMnTrueS()==false){return false};
  if(IsTableInputMaTrueS()==false){return false};
  if(IsTableInputMpTrueS()==false){return false};
  RemoveRowHiddenS();
  if(bodyBillPur.childElementCount==0){PurchasesItems.style.border="2px solid #ff0000";return false}else{PurchasesItems.style.border="2px solid rgb(155, 153, 153)";}
  return true;
}

function RemoveRowHiddenS(){
  let ArraryNamesRow=[];
  let ArraryNames=[];
    let PRBRow=document.getElementsByClassName("PRBRowS");
    for (let index1 = PRBRow.length - 1; index1 >= 0 ; index1--) {
      if (PRBRow.item(index1).hidden==true){
        PRBRow.item(index1).remove();
      }
    }
    for (let index = 0; index < PRBRow.length; index++) {
        var RowN=PRBRow.item(index).children;
        for (let indexN = 0; indexN < RowN.length; indexN++) {
          var RowNi =RowN.item(indexN).children.item(0)
          if(RowNi.name !=""){
            ArraryNames.push(RowNi.name);
          }
        }
        ArraryNamesRow.push(ArraryNames);
        ArraryNames=[];
    }
    document.getElementById("ArrayTableS").value=ArraryNamesRow;
}

function IsTableInputMpTrueS(){
  let TableInputMp=document.getElementsByClassName("TableInputMpS");
  for (let index = 0; index < TableInputMp.length; index++) {
    if (TableInputMp.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMp.item(index).value==""){
        TableInputMp.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMp.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}

function IsTableInputMaTrueS(){
  let TableInputMa=document.getElementsByClassName("TableInputMaS");
  for (let index = 0; index < TableInputMa.length; index++) {
    if (TableInputMa.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMa.item(index).value==""){
        TableInputMa.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMa.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}

function IsTableInputMcTrueS(){
  let TableInputMc=document.getElementsByClassName("TableInputMcS");
  for (let index = 0; index < TableInputMc.length; index++) {
    if (TableInputMc.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMc.item(index).value==""){
        TableInputMc.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMc.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}
function IsTableInputMnTrueS(){
  let TableInputMn=document.getElementsByClassName("TableInputMnS");
  for (let index = 0; index < TableInputMn.length; index++) {
    if (TableInputMn.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMn.item(index).value==""){
        TableInputMn.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMn.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}
function onsubmitFormS1(){
  if(IstrueDataInformS()===true){
    document.getElementById("BillNumberS").value=MaxBillNumberS();
    document.getElementById("ModeS").value="1";
    onsubmitFormS(6000);
  }
}

function BillNumberIsfoundS(BillNumber){
  for (let index = 0; index < DataSales.length; index++) {
    if(DataSales[index].BillNumber==BillNumber ){
      return true;
    }
}
return false;
}

function onsubmitFormS2(){
  let BillNumberS =document.getElementById("BillNumberS");
  if(IstrueDataInformS()===true){
    if(BillNumberIsfoundS(BillNumberS.value)==true){
      document.getElementById("ModeS").value="2";
      onsubmitFormS(10000);
    }
  }
}

function onsubmitFormS3(){
  let BillNumberS =document.getElementById("BillNumberS");
  if(BillNumberIsfoundS(BillNumberS.value)==true){
  document.getElementById("ModeS").value="3";
  onsubmitFormS(8000);
  }
}

function onsubmitFormS(Time){
  document.getElementById("typeS").value="2";
  document.getElementById("UserNameS").value=localStorage.getItem("User_Name");
  let MainForm=document.getElementById("FormSalesDetails");
  var w = window.open('', 'form_target', 'width=600, height=400');
  MainForm.target = 'form_target';
  MainForm.action='https://script.google.com/macros/s/AKfycbzrvVLAeb6YwTxzwZq6nSLuerWq4qnUFuY2_7o2NO76gqHAlna6NC9wmBEXBFiSrzjScw/exec'
  MainForm.submit();
  if (MainForm.onsubmit()==true){
    const myTimeout = setTimeout(function(){ 
                w.close();
                clearTimeout(myTimeout)
                location.reload();
    }, Time);
  }
} 


// **************************SalesBrowser***********

function LoadSalesToTable(){
  const myTimeout = setTimeout(function(){ 
  document.getElementById("bodydataS").innerHTML=""
  if (isNaN(DataSales[0].Num)){return}
  for (let index = 0; index < DataSales.length; index++) {
    if(DataSales[index].Num!="" ){
      AddRowPrS(DataSales[index].BillNumber,DataSales[index].EDate,DataSales[index].AccountCode,DataSales[index].Total,DataSales[index].Note,DataSales[index].UserName)
    }
  }
  clearTimeout(myTimeout)
}, 1000);
}

function AddRowPrS(BillNumberP,BillDateP,AccountCodeP,TotalP,NoteP,UserNameP) {
  let bodydata=document.getElementById("bodydataS");
  let row = bodydata.insertRow();
  row.id="S" + bodydata.childElementCount;
  let cell = row.insertCell();
  cell.id="S" + bodydata.childElementCount + "BillNumberS";
  cell.innerHTML = BillNumberP;
  cell = row.insertCell();
  cell.id="S" + bodydata.childElementCount + "BillDateS";
  cell.innerHTML = GetDateFromString(BillDateP);
  cell = row.insertCell();
  cell.id="S" + bodydata.childElementCount + "AccountCodeS";
  cell.innerHTML = AccountCodeP;
  cell = row.insertCell();
  cell.id="S" + bodydata.childElementCount + "AccountNameS";
  cell.innerHTML = GetNameAccount(AccountCodeP);
  cell = row.insertCell();
  cell.id="S" + bodydata.childElementCount + "TotalS";
  cell.innerHTML = TotalP;
  cell = row.insertCell();
  cell.id="S" + bodydata.childElementCount + "NoteS";
  cell.innerHTML = NoteP;
  cell = row.insertCell();
  cell.id="S" + bodydata.childElementCount + "UserNameS";
  cell.innerHTML = UserNameP;
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "button";
  btb.id="ButS" + bodydata.childElementCount;
  btb.value = "تعديل";
  td.appendChild(btb)
  btb.style.cursor="pointer";
  btb.style.color="red";
  btb.style.padding="2px";
  btb.onclick=function(){showdatarowsS()};
  };

function MaxBillNumberS(){
  let XX;
  let BillNumbers=[];
  for (let index = 0; index < DataSales.length; index++) {
    if(DataSales[index].Num!="" ){
      BillNumbers.push(DataSales[index].BillNumber);
    }
  }
  XX= Math.max.apply(Math, BillNumbers) + 1;
  if(isNaN(XX)==true){return 1}else{return XX}
}
  function ShowDetailsS(){
    ShowSelectForm("SalesWi");
    document.getElementById("BillNumberS").value=MaxBillNumberS();
    document.getElementById("BillDateS").value=""
    document.getElementById("AccountCodeS").value=""
    document.getElementById("AccountNameS").value=""
    document.getElementById("NoteS").value=""
    document.getElementById("STotal").value=""
    document.getElementById("Discount").value=""
    document.getElementById("STotalNet").value=""
    document.getElementById("bodyBillSa").innerHTML=""    
  }

  function showdatarowsS() {
    let indextable= document.activeElement.parentElement.parentElement.id;
    let IndexRow =document.getElementById(indextable).children.item(0).textContent  ;
    let NameAccountTable=document.getElementById(indextable).children.item(3).textContent  ;
    ShowDetailsS();
    for (let index = 0; index < DataSales.length; index++) {
      if(DataSales[index].BillNumber == IndexRow){
        document.getElementById("BillNumberS").value= IndexRow;
        document.getElementById("BillDateS").value=GetDateFromString(DataSales[index].EDate);
        document.getElementById("AccountCodeS").value=DataSales[index].AccountCode;
        document.getElementById("AccountNameS").value=NameAccountTable;
        document.getElementById("NoteS").value=DataSales[index].Note;
        document.getElementById("ShipS").value=DataSales[index].Ship;
        document.getElementById("Discount").value=DataSales[index].Discount;
        document.getElementById("DiscountPaid").value=DataSales[index].DiscountPaid;
        document.getElementById("STotalNet").value=DataSales[index].Total;
      }
    }
    let IndexTableRow = 1;
    for (let indexM = 0; indexM < DataMove.length; indexM++) {
      if(DataMove[indexM].BillNumber == IndexRow && DataMove[indexM].Type=="2"){
        AddRowPrElementS();
        document.getElementById(`PRBCodeS${IndexTableRow}`).value= DataMove[indexM].MatCode;
        document.getElementById(`PRBCodeNS${IndexTableRow}`).value=GetMatNameFormCode1(DataMove[indexM].MatCode)
        document.getElementById(`PRBAmountS${IndexTableRow}`).value= Math.abs(DataMove[indexM].Amount);
        document.getElementById(`PRBWareHS${IndexTableRow}`).value= DataMove[indexM].WareHouse;
        document.getElementById(`PRBPriceS${IndexTableRow}`).value= DataMove[indexM].Price;
        document.getElementById(`PRBMatTotalS${IndexTableRow}`).value= DataMove[indexM].MatTotal;
        document.getElementById(`PRBNoteS${IndexTableRow}`).value= DataMove[indexM].Note;
        IndexTableRow++;
      }
    }
    CaluclateTotalS();
    document.getElementById("CountRowS").value=document.getElementById("CountRowS1").value
  };


// ********************EntryWi
function LoadEntry(){
  DataEntry=[];
  fetch(UrlEntry)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataEntry.push(rowUser);
      })
  })
}
function LoadTeacher(){
  DataTeacher=[];
  fetch(UrlTeacher)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataTeacher.push(rowUser);
      })
  })
} 
function LoadAccoutName(Num){
  let MatName,MatCode;
  let optionClass;
  let listMatName =document.getElementById("listMatNameE" + Num);
  listMatName.innerHTML="";
  for (let index = 0; index < DataAccounts.length; index++) {
    MatCode=DataAccounts[index].AccountCode
    MatName=DataAccounts[index].AccountName
    if(MatCode!=""){
      optionClass=document.createElement("option");
      optionClass.value=MatName;
      optionClass.textContent=MatName;
      listMatName.appendChild(optionClass);
    }
  }
}

function GetCodeMatFormNameE(NameA, TRow){
  let Row=String(TRow).slice(9,String(TRow).length)
  for (let index = 0; index < DataAccounts.length; index++) {
    if(DataAccounts[index].AccountName==NameA ){
      document.getElementById("PRBCodeE" + Row).value=DataAccounts[index].AccountCode;
      return ;
    }
  }
  document.getElementById("PRBCodeE" + Row).value="";
}

function GetNameMatFormCodeE(Code , TRow){
  let Row=String(TRow).slice(8,String(TRow).length);
  for (let index = 0; index < DataAccounts.length; index++) {
    if(DataAccounts[index].AccountCode==Code ){
      document.getElementById("PRBCodeNE" + Row).value=DataAccounts[index].AccountName;
      return ;
    }
  }
  document.getElementById("PRBCodeNE" + Row).value="";
}

function CaluclateTotalE(){
  let PTotal = document.getElementById("ETotal");
  let CountRow= document.getElementById("CountRowE1");
  let TableInputMt=document.getElementsByClassName("TableInputMpE");
  let total = 0 ;
  let Rows = 0 ;
  for (let index = 0; index < TableInputMt.length; index++) {
    if (TableInputMt.item(index).parentElement.parentElement.hidden==false){
      total =  total + Number(TableInputMt.item(index).value);
      Rows= Rows + 1
    }
  }
  PTotal.value=total;
  CountRow.value=Rows;
}

function AddRowPrElementE() {
  var td;
  let bodydata=document.getElementById("bodyBillEn");
  let row = bodydata.insertRow();
  row.id="PRBE" + bodydata.childElementCount;
  row.className="PRBRowE";
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "number";
  btb.autocomplete="off";
  btb.className="TableInputMpE";
  btb.style.width="92%";
  btb.id="PRBTotalE" + bodydata.childElementCount ;
  btb.name="Total" + bodydata.childElementCount ;
  btb.onkeyup=function(){CaluclateTotalE()};
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "text";
  btb.autocomplete="off";
  btb.className="TableInputMcE";
  btb.style.width="89%";
  btb.id="PRBCodeE" + bodydata.childElementCount ;
  btb.name="MatCode" + bodydata.childElementCount ;
  btb.onkeyup=function(){GetNameMatFormCodeE(this.value,this.id)};
  td.appendChild(btb)
  row.appendChild(td=document.createElement('td'));
  td.innerHTML=`<input list='listMatNameE${bodydata.childElementCount}' id="PRBCodeNE${bodydata.childElementCount}"  style='width:95%;' class="TableInputMnE" onchange="GetCodeMatFormNameE(this.value,this.id)" autocomplete="off">`;
  var datalist = document.createElement('datalist');
  datalist.id=`listMatNameE${bodydata.childElementCount}`;
  td.appendChild(datalist);
  LoadAccoutName(bodydata.childElementCount);
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "text";
  btb.autocomplete="off";
  btb.className="TableInputMnoE";
  btb.style.width="92%";
  btb.id="PRBNoteE" + bodydata.childElementCount ;
  btb.name="Note" + bodydata.childElementCount ;
  td.appendChild(btb);
  row.appendChild(td=document.createElement('td'));
  btb = document.createElement('input');
  btb.type = "button";
  btb.id="PRBDS" + bodydata.childElementCount;
  btb.value = "X";
  btb.className="BtnStyle";
  btb.onclick=function(){DeleteRowE(this.id)};
  td.appendChild(btb)
  };


  function DeleteRowE(TRow){
    let Row=String(TRow).slice(5,String(TRow).length);
    document.getElementById("PRBE" + Row).hidden=true;
    CaluclateTotalE();
  }
  
function IstrueDataInformE(){
  let BillDateP=document.getElementById("BillDateE");
  let bodyBillPur=document.getElementById("bodyBillEn"); 
  let PurchasesItems=document.getElementById("EntryItems"); 
  let ETotal  =document.getElementById("ETotal"); 
  if(BillDateP.value==""){BillDateP.style.border="2px solid #ff0000";return false}else{BillDateP.style.border="none";}
  if(bodyBillPur.childElementCount==0){PurchasesItems.style.border="2px solid #ff0000";return false}else{PurchasesItems.style.border="2px solid rgb(155, 153, 153)";}
  if(ETotal.value!=0){ETotal.style.border="2px solid #ff0000";return false}else{ETotal.style.border="2px solid rgb(155, 153, 153)";}
  if(IsTableInputMpTrueE()==false){return false};
  if(IsTableInputMcTrueE()==false){return false};
  if(IsTableInputMnTrueE()==false){return false};
  if(IsTableInputMaTrueE()==false){return false};
  RemoveRowHiddenE();
  if(ETotal.value!=0){ETotal.style.border="2px solid #ff0000";return false}else{ETotal.style.border="2px solid rgb(155, 153, 153)";}
  
  return true;
}

function RemoveRowHiddenE(){
  let ArraryNamesRow=[];
  let ArraryNames=[];
    let PRBRow=document.getElementsByClassName("PRBRowE");
    for (let index1 = PRBRow.length - 1; index1 >= 0; index1--) {
      if (PRBRow.item(index1).hidden==true){
        PRBRow.item(index1).remove();
      }
    }
    for (let index = 0; index < PRBRow.length; index++) {
        var RowN=PRBRow.item(index).children;
        for (let indexN = 0; indexN < RowN.length; indexN++) {
          var RowNi =RowN.item(indexN).children.item(0)
          if(RowNi.name !=""){
            ArraryNames.push(RowNi.name);
          }
        }
        ArraryNamesRow.push(ArraryNames);
        ArraryNames=[];
    }
    document.getElementById("ArrayTableE").value=ArraryNamesRow;
}

function IsTableInputMpTrueE(){
  let TableInputMp=document.getElementsByClassName("TableInputMpE");
  for (let index = 0; index < TableInputMp.length; index++) {
    if (TableInputMp.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMp.item(index).value==""){
        TableInputMp.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMp.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}

function IsTableInputMaTrueE(){
  let TableInputMa=document.getElementsByClassName("TableInputMnoE");
  for (let index = 0; index < TableInputMa.length; index++) {
    if (TableInputMa.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMa.item(index).value==""){
        TableInputMa.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMa.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}

function IsTableInputMcTrueE(){
  let TableInputMc=document.getElementsByClassName("TableInputMcE");
  for (let index = 0; index < TableInputMc.length; index++) {
    if (TableInputMc.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMc.item(index).value==""){
        TableInputMc.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMc.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}
function IsTableInputMnTrueE(){
  let TableInputMn=document.getElementsByClassName("TableInputMnE");
  for (let index = 0; index < TableInputMn.length; index++) {
    if (TableInputMn.item(index).parentElement.parentElement.hidden==false){
      if (TableInputMn.item(index).value==""){
        TableInputMn.item(index).parentElement.style.border="2px solid #ff0000";
        return false ;
      }else{
        TableInputMn.item(index).parentElement.style.border="2px solid rgb(155, 153, 153)";
      }
    }
  }
  return true ;
}
function onsubmitFormE1(){
  if(IstrueDataInformE()===true){
    document.getElementById("BillNumberE").value=MaxBillNumberE();
    document.getElementById("ModeE").value="1";
    onsubmitFormE(5000);
  }
}
function BillNumberIsfoundE(BillNumber){
  for (let index = 0; index < DataEntry.length; index++) {
    if(DataEntry[index].BillNumber==BillNumber ){
      return true;
    }
}
return false;
}

function onsubmitFormE2(){
  let BillNumberE =document.getElementById("BillNumberE");
  if(IstrueDataInformE()===true){
    if(BillNumberIsfoundE(BillNumberE.value)==true){
      document.getElementById("ModeE").value="2";
      onsubmitFormE(12000);
    }
  }
}

function onsubmitFormE3(){
  let BillNumberE =document.getElementById("BillNumberE");
  if(BillNumberIsfoundE(BillNumberE.value)==true){
    document.getElementById("ModeE").value="3";
    onsubmitFormE(8000);
  }
}

function onsubmitFormE(Time){
  document.getElementById("typeE").value="3";
  document.getElementById("UserNameE").value=localStorage.getItem("User_Name");
  let MainForm=document.getElementById("FormEntryDetails");
  var w = window.open('', 'form_target', 'width=600, height=400');
  MainForm.target = 'form_target';
  MainForm.action='https://script.google.com/macros/s/AKfycbzrvVLAeb6YwTxzwZq6nSLuerWq4qnUFuY2_7o2NO76gqHAlna6NC9wmBEXBFiSrzjScw/exec'
  MainForm.submit();
  if (MainForm.onsubmit()==true){
    const myTimeout = setTimeout(function(){ 
                w.close();
                clearTimeout(myTimeout)
                location.reload();
    }, Time);
  }
} 


// **************************EntryBrowser***********
function LoadEntryToTable(){
  const myTimeout = setTimeout(function(){ 
  document.getElementById("bodydataE").innerHTML=""
  if (isNaN(DataEntry[0].Num)){return}
  for (let index = 0; index < DataEntry.length; index++) {
    if(DataEntry[index].Num!="" ){
      AddRowPrE(DataEntry[index].BillNumber,DataEntry[index].EDate,DataEntry[index].Note,DataEntry[index].UserName)
    }
  }
  clearTimeout(myTimeout)
}, 1000);
}

function AddRowPrE(BillNumberP,BillDateP,NoteP,UserNameP) {
  let bodydata=document.getElementById("bodydataE");
  let row = bodydata.insertRow();
  row.id="E" + bodydata.childElementCount;
  let cell = row.insertCell();
  cell.id="E" + bodydata.childElementCount + "BillNumberE";
  cell.innerHTML = BillNumberP;
  cell = row.insertCell();
  cell.id="E" + bodydata.childElementCount + "BillDateE";
  cell.innerHTML = GetDateFromString(BillDateP);
  cell = row.insertCell();
  cell.id="E" + bodydata.childElementCount + "NoteE";
  cell.innerHTML = NoteP;
  cell = row.insertCell();
  cell.id="E" + bodydata.childElementCount + "UserNameE";
  cell.innerHTML = UserNameP;
  row.appendChild(td=document.createElement('td'));
  var btb = document.createElement('input');
  btb.type = "button";
  btb.id="ButE" + bodydata.childElementCount;
  btb.value = "تعديل";
  td.appendChild(btb)
  btb.style.cursor="pointer";
  btb.style.color="red";
  btb.style.padding="2px";
  btb.onclick=function(){showdatarowsE()};
  };

function MaxBillNumberE(){
  let XX;
  let BillNumbers=[];
  for (let index = 0; index < DataEntry.length; index++) {
    if(DataEntry[index].Num!="" ){
      BillNumbers.push(DataEntry[index].BillNumber);
    }
  }
  XX= Math.max.apply(Math, BillNumbers) + 1;
  if(isNaN(XX)==true){return 1}else{return XX}
}
  function ShowDetailsE(){
    ShowSelectForm("EntryWi");
    document.getElementById("BillNumberE").value=MaxBillNumberE();
    document.getElementById("BillDateE").value=""
    document.getElementById("NoteE").value=""
    document.getElementById("ETotal").value=""
    document.getElementById("bodyBillEn").innerHTML=""    
  }

  function showdatarowsE() {
    let indextable= document.activeElement.parentElement.parentElement.id;
    let IndexRow =document.getElementById(indextable).children.item(0).textContent  ;
    ShowDetailsE();
    for (let index = 0; index < DataEntry.length; index++) {
      if(DataEntry[index].BillNumber == IndexRow){
        document.getElementById("BillNumberE").value= IndexRow;
        document.getElementById("BillDateE").value=GetDateFromString(DataEntry[index].EDate);
        document.getElementById("NoteE").value=DataEntry[index].Note;
      }
    }
    let IndexTableRow = 1;
    for (let indexM = 0; indexM < DataTeacher.length; indexM++) {
      if(DataTeacher[indexM].BillNumber == IndexRow && DataTeacher[indexM].Type=="3"){
        AddRowPrElementE();
        document.getElementById(`PRBTotalE${IndexTableRow}`).value= DataTeacher[indexM].Total;
        document.getElementById(`PRBCodeE${IndexTableRow}`).value= DataTeacher[indexM].AccountCode;
        document.getElementById(`PRBCodeNE${IndexTableRow}`).value=GetNameAccount(DataTeacher[indexM].AccountCode)
        document.getElementById(`PRBNoteE${IndexTableRow}`).value= DataTeacher[indexM].Note;
        IndexTableRow++;
      }
    }
    CaluclateTotalE();
    document.getElementById("CountRowE").value=document.getElementById("CountRowE1").value
  };


// **************************AccountsBrowser***********
function LoadAccountsSeting(){
  DataAccountsSeting=[];
  fetch(UrlAccountsSeting)
  .then(res => res.text())
  .then(rep => {
      const jsonData = JSON.parse(rep.substring(47).slice(0, -2));
      const colzUser = [];
      jsonData.table.cols.forEach((heading) => {
          if (heading.label) {
              let columnUser = heading.label;
              colzUser.push(columnUser);
          }
      })
      jsonData.table.rows.forEach((rowData) => {
          const rowUser = {};
          colzUser.forEach((ele, ind) => {
              rowUser[ele] = (rowData.c[ind] != null) ? rowData.c[ind].v : '';
          })
          DataAccountsSeting.push(rowUser);
      })
      LoadSelectAccounts();
  })
}
function LoadSelectAccounts(){
  let optionClass;
  let BelongeTo =document.getElementById("BelongeTo");
  let AccountClass =document.getElementById("AccountClass");
  BelongeTo.innerHTML=""
  AccountClass.innerHTML=""
  optionClass=document.createElement("option");
  optionClass.value="All";
  optionClass.textContent="All";
  AccountClass.appendChild(optionClass);
  optionClass=document.createElement("option");
  optionClass.value="All";
  optionClass.textContent="All";
  BelongeTo.appendChild(optionClass);
  for (let index = 0; index < DataAccountsSeting.length; index++) {
    var XX=DataAccountsSeting[index].AccountClass;
    var YY=DataAccountsSeting[index].BelongeTo;
    if(DataAccountsSeting[index].Num!=""){
    if(XX!="" && isNaN(XX)==true){
      optionClass=document.createElement("option");
      optionClass.value=XX;
      optionClass.textContent=XX;
      AccountClass.appendChild(optionClass);
    }
    if(YY!="" && isNaN(YY)==true){
    optionClass=document.createElement("option");
    optionClass.value=YY;
    optionClass.textContent=YY;
    BelongeTo.appendChild(optionClass);
    }
  }
  }
  
  document.getElementById("DiscountPaid").value=DataAccountsSeting[2].BelongeTo
  document.getElementById("ShipS").value=DataAccountsSeting[3].BelongeTo
  document.getElementById("ShipPr").value=DataAccountsSeting[3].BelongeTo
}

function LoadAccountsToTable(){
  let BelongeTo =document.getElementById("BelongeTo");
  let AccountClass =document.getElementById("AccountClass");
  let TableSearch =document.getElementById("TableSearch");
  const myTimeout = setTimeout(function(){ 
  document.getElementById("bodyTable").innerHTML=""
  for (let index = 0; index < DataAccounts.length; index++) {
    if(DataAccounts[index].AccountNum!="" ){
      if(BelongeTo.value=="All" && AccountClass.value=="All" && TableSearch.value==""){
      AddRowAccounts(DataAccounts[index].AccountNum,DataAccounts[index].AccountCode,DataAccounts[index].AccountName,DataAccounts[index].BelongeTo,DataAccounts[index].AccountClass,DataAccounts[index].AccountBalance)
      }else if(BelongeTo.value==DataAccounts[index].BelongeTo && AccountClass.value=="All" && TableSearch.value==""){
        AddRowAccounts(DataAccounts[index].AccountNum,DataAccounts[index].AccountCode,DataAccounts[index].AccountName,DataAccounts[index].BelongeTo,DataAccounts[index].AccountClass,DataAccounts[index].AccountBalance)
      }else if(BelongeTo.value=="All" && AccountClass.value==DataAccounts[index].AccountClass && TableSearch.value==""){
      AddRowAccounts(DataAccounts[index].AccountNum,DataAccounts[index].AccountCode,DataAccounts[index].AccountName,DataAccounts[index].BelongeTo,DataAccounts[index].AccountClass,DataAccounts[index].AccountBalance)
      }else if(BelongeTo.value=="All" && AccountClass.value=="All" && TableSearch.value==DataAccounts[index].AccountName){
        AddRowAccounts(DataAccounts[index].AccountNum,DataAccounts[index].AccountCode,DataAccounts[index].AccountName,DataAccounts[index].BelongeTo,DataAccounts[index].AccountClass,DataAccounts[index].AccountBalance)
      }else if(BelongeTo.value=="All" && AccountClass.value==DataAccounts[index].AccountClass && TableSearch.value==DataAccounts[index].AccountName){
        AddRowAccounts(DataAccounts[index].AccountNum,DataAccounts[index].AccountCode,DataAccounts[index].AccountName,DataAccounts[index].BelongeTo,DataAccounts[index].AccountClass,DataAccounts[index].AccountBalance)
      }else if(BelongeTo.value==DataAccounts[index].BelongeTo && AccountClass.value=="All" && TableSearch.value==DataAccounts[index].AccountName){
        AddRowAccounts(DataAccounts[index].AccountNum,DataAccounts[index].AccountCode,DataAccounts[index].AccountName,DataAccounts[index].BelongeTo,DataAccounts[index].AccountClass,DataAccounts[index].AccountBalance)
      }else if(BelongeTo.value==DataAccounts[index].BelongeTo && AccountClass.value==DataAccounts[index].AccountClass && TableSearch.value==DataAccounts[index].AccountName){
        AddRowAccounts(DataAccounts[index].AccountNum,DataAccounts[index].AccountCode,DataAccounts[index].AccountName,DataAccounts[index].BelongeTo,DataAccounts[index].AccountClass,DataAccounts[index].AccountBalance)
      }else if(BelongeTo.value==DataAccounts[index].BelongeTo && AccountClass.value==DataAccounts[index].AccountClass && TableSearch.value==""){
        AddRowAccounts(DataAccounts[index].AccountNum,DataAccounts[index].AccountCode,DataAccounts[index].AccountName,DataAccounts[index].BelongeTo,DataAccounts[index].AccountClass,DataAccounts[index].AccountBalance)
      }
  }
  }
  clearTimeout(myTimeout)
}, 1000);
}

function AddRowAccounts(AccountNum,AccountCode,AccountName,BelongeTo,AccountClass,AccountBalance) {
  let bodydata=document.getElementById("bodyTable");
  let row = bodydata.insertRow();
  row.id="ACC" + bodydata.childElementCount;
  let cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "AccountNum";
  cell.innerHTML = AccountNum;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "AccountCode";
  cell.innerHTML =AccountCode;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "AccountName";
  cell.innerHTML = AccountName;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "BelongeTo";
  cell.innerHTML = BelongeTo;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "AccountClass";
  cell.innerHTML = AccountClass;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "AccountBalance";
  cell.innerHTML = AccountBalance;
  };

// **************************MatsBrowser***********
function LoadMatsToTable(){
  let TableSearchMat =document.getElementById("TableSearchMat");
  const myTimeout = setTimeout(function(){ 
  document.getElementById("bodyTableMats").innerHTML=""
  for (let index = 0; index < DataMat.length; index++) {
    if(DataMat[index].MatNum!=""){
      if(TableSearchMat.value==""){
      AddRowMats(DataMat[index].MatNum,DataMat[index].MatCode,DataMat[index].MatName,DataMat[index].MatBalance,DataMat[index].MatCost,DataMat[index].Value)
      }else{
    if(TableSearchMat.value==DataMat[index].MatName ){
      AddRowMats(DataMat[index].MatNum,DataMat[index].MatCode,DataMat[index].MatName,DataMat[index].MatBalance,DataMat[index].MatCost,DataMat[index].Value)
      }
    }
  }
  }
  clearTimeout(myTimeout)
}, 1000);
}

function AddRowMats(MatNum,MatCode,MatName,MatBalance,MatCost,Value) {
  let bodydata=document.getElementById("bodyTableMats");
  let row = bodydata.insertRow();
  row.id="ACC" + bodydata.childElementCount;
  let cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "MatNum";
  cell.innerHTML = MatNum;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "MatCode";
  cell.innerHTML =MatCode;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "MatName";
  cell.innerHTML = MatName;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "MatBalance";
  cell.innerHTML = MatBalance;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "MatCost";
  cell.innerHTML = MatCost;
  cell = row.insertCell();
  cell.id="ACC" + bodydata.childElementCount + "Value";
  cell.innerHTML = Value;
  };



// ***********************Mode*********************
function ConvertMode(){
  if (localStorage.getItem("FColor")==1){
    ConvertModeToSun();
  }else{
    ConvertModeToMoon();
  }
 }

function ConvertModeToSun(){
  localStorage.setItem("FColor", 1);
  document.getElementById("Moon").style.display="inline-block";
  document.getElementById("Sun").style.display="none";
  document.querySelector(':root').style.setProperty('--FColor', "wheat"); 
  document.querySelector(':root').style.setProperty('--EColor', "white");
  document.querySelector(':root').style.setProperty('--loginColor', "whitesmoke"); 
  document.querySelector(':root').style.setProperty('--FontColor', "#f2a20b"); 
  document.querySelector(':root').style.setProperty('--Font2Color', "#a53333"); 
  document.querySelector(':root').style.setProperty('--Font3Color', "#a53333");
  document.querySelector(':root').style.setProperty('--THColor', "wheat");  
  document.querySelector(':root').style.setProperty('--TDColor', "yellow"); 
} 
function ConvertModeToMoon(){
  localStorage.setItem("FColor", 2);
  document.getElementById("Sun").style.display="inline-block";
  document.getElementById("Moon").style.display="none";
  document.querySelector(':root').style.setProperty('--FColor', "#141e30"); 
  document.querySelector(':root').style.setProperty('--EColor', "#243b55");
  document.querySelector(':root').style.setProperty('--loginColor', "#00000080"); 
  document.querySelector(':root').style.setProperty('--FontColor', "white"); 
  document.querySelector(':root').style.setProperty('--Font2Color', "#d3f6f8"); 
  document.querySelector(':root').style.setProperty('--Font3Color', "black"); 
  document.querySelector(':root').style.setProperty('--THColor', "gray");  
  document.querySelector(':root').style.setProperty('--TDColor', "Red"); 
}  
