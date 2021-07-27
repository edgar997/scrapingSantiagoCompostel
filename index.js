require('chromedriver');
const colors = require('colors')
const {isNull, isEmpty,isError,isNaN, isUndefined, indexOf} = require('lodash');
const {Builder, By, Key, until,Capabilities} = require('selenium-webdriver');
const caps = new Capabilities();
const fs = require('fs');
var excel = require("exceljs");
var workbook1 = new excel.Workbook();

caps.setPageLoadStrategy("eager");

let listaDatos=[],
url_actual = 'https://www.usc.es/gl/directorio?q=&l=a&c=0&p=1';
async function urlActual(driver){
  url_actual = await driver.getCurrentUrl();
  console.log(colors.green("nueva url= "+url_actual))
} 
async function avanzarNum(driver){
  const actions = driver.actions({async: true});
  let err = false;
  let btn = await driver.findElement(By.css('.pagination-list .pagination-list-item .is-next'))
  .then((r)=>{
    return r
  },(e)=>{
    console.log(colors.magenta('errors'))
    err = true;
    
  });

  if(err){
    let num= await avanzarLetra(driver)
    console.log(colors.grey('numero de letre: '+num))
    if(num == 26){
      url_actual="";
      return;
    }else{
      console.log(colors.green('cambiando a letras'))
      await urlActual(driver)
      return;
      
    }
  }
  let clase =await btn.getAttribute("class")
  clase = clase.split(" ");
 // console.log(clase)
    if(clase.indexOf('is-current')!==-1){
      //si existe is-current
      let num=  await avanzarLetra(driver)
      console.log(colors.grey('numero de letre: '+num))
      if(num == 26){
        url_actual="";
      }else{
        console.log(colors.green('cambiando a letras'))
        await urlActual(driver)

      }


    }else{
      console.log(colors.green('haciendo click numeros'))
      await actions.move({origin:btn}).press().perform();
      await actions.move({origin:btn}).release().perform();
      //await btn.click();
      await urlActual(driver)
    }

}

async function avanzarLetra(driver){
  const actions = driver.actions({async: true});
  let contador = 0;
  let listaLetra = await driver.findElements(By.css('.alphabetic-filter-letters-list li'));
  //listaLetra
  
  for (let e of listaLetra) {
    let v = await e.findElement(By.css('.at-alphabetic-filter-letter'));
    console.log(await v.getAttribute('class'))
    v = await v.getAttribute('class')
    v = v.split(" ")
    
    if (v.indexOf('is-selected')!==-1){
      let mostrar = contador +1
      console.log(colors.green('valor de contador: '+mostrar))      
      //await listaLetra[`${contador+1}`].click();
     // console.log(await listaLetra[mostrar].getAttribute("class"))
      await actions.move({origin:listaLetra[`${contador+1}`]}).press().perform();
      await actions.move({origin:listaLetra[`${contador+1}`]}).release().perform();
      return mostrar; 
      contador = 0;
    }else{
      contador = contador +1;
    }
  }
  


}
async function scaner(driver,url){
  console.log(colors.green('iniciado scaneo'));
  await driver.get(url);

  let constent  = await driver.findElement(By.css('.repository-content'));
    let elements =   await driver.wait(constent.findElements(By.css('.col-md-6')));
    for(let e of elements) {
      let name = await e.findElement(By.css('.at-title')).getText();
      let correo,telefono,area,campus,tlf=false,tlfv;
      await e.findElement(By.css('.rot13')).getText()
      .then((r)=>{correo= r},(e)=>{correo = "N/P";})
      
      let dl = await e.findElement(By.css('.at-desc-list'))
      let dt = await dl.findElements(By.css('.at-desc-list dt'))
      let dd = await dl.findElements(By.css('.at-desc-list dd'))
      let cont = 0;
      for(let i of dt){
       //console.log(await i.getText())
        if(await i.getText()=="Campus"){
          await dd[cont].getText().then((r)=>{campus = r},(e)=>{campus = "N/P"})
        }else{
          campus="N/P"
        }

        if(await i.getText() == "Área"){  
          area = await dd[cont].getText()
          .then((r)=>{return r},(e)=>{ console.log(e); return "N/P"});
        }else{
          area="N/P"
        }

        if(await i.getText()==="Teléfono"){
          telefono = await dd[cont].getText()
          .then((r)=>{ tlfv=r; return r;});
          tlf =true;
         }else{
           
           telefono="N/P"
         }
        cont = cont+1;
      }

      listaDatos.push(
         {nombre: `${name}`,
         correo: `${ correo }`,
         telefono : `${(tlf)? telefono = tlfv: telefono}`,
         area: `${area}`,
         campus: `${campus}`
        });
      }
      console.log(listaDatos.length)
} 
async function execute() {
  let err=false;
  let driver = await new Builder().withCapabilities(caps).forBrowser('chrome').build();
  await driver.manage().window().setRect({ width: 1024, height: 730 });
  try {
    

    
    do {
        console.log(colors.red(url_actual));
        await scaner(driver,url_actual)
        await avanzarNum(driver)
    } while (url_actual!=="" );

    workbook1.creator = 'edgar parra';
    workbook1.lastModifiedBy = 'edgar parra';
    workbook1.created = new Date();
    workbook1.modified = new Date();
    var sheet1 = workbook1.addWorksheet('libro1');
    //console.log(listaDatos)
    console.log(listaDatos.length)
    var reColumns=[
      {header:'nombre',key:'nombre'},
      {header:'correo',key:'correo'},
      {header:'telefono',key:'telefono'},
      {header:'area',key:'area'},
      {header:'campus',key:'campus'},

  ];
  sheet1.columns = reColumns;
  sheet1.addRows(listaDatos)
  workbook1.xlsx.writeFile("error.xlsx").then(function() {
      console.log("xlsx file is written.");
  });
    console.log(colors.green('en un momento se iniciara la descarga'))
    
    //await driver.quit();
    
  }catch(error){
    err = true
    console.log(colors.red(error))
    /*workbook1.creator = 'edgar parra';
    workbook1.lastModifiedBy = 'edgar parra';
    workbook1.created = new Date();
    workbook1.modified = new Date();
    var sheet1 = workbook1.addWorksheet('libro1');
    //console.log(listaDatos)
    console.log(listaDatos.length)
    var reColumns=[
      {header:'nombre',key:'nombre'},
      {header:'correo',key:'correo'},
      {header:'telefono',key:'telefono'},
      {header:'area',key:'area'},
      {header:'campus',key:'campus'},

  ];
  sheet1.columns = reColumns;
  sheet1.addRows(listaDatos)
  workbook1.xlsx.writeFile("error.xlsx").then(function() {
      console.log("xlsx file is written.");
  });*/
    
  } 
  if(err){
    err = false
    console.log(listaDatos.length)

    await execute()
  }
}

 execute()