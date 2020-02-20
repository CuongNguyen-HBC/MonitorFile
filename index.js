const chokidar = require('chokidar')
const process = require('child_process')
const path = require('path')
const sql = require('mssql')
const fs = require('fs')
const nodemail = require('nodemailer')
const config =  {
   user: 'sa',
   password: '*hbc123',
   server: '118.69.225.159', // You can use 'localhost\\instance' to connect to named instance
   database: 'HBC',
   port:8500
}
const filexml = []
filexml['HBC']    = 'C:\\Users\\cuong.nguyen\\Desktop\\Exec\\HBC.xml'
filexml['KBI']    = 'C:\\Users\\cuong.nguyen\\Desktop\\Exec\\KBI.xml'
filexml['LBC']    = 'C:\\Users\\cuong.nguyen\\Desktop\\Exec\\LBC.xml'
filexml['AKBC']   = 'C:\\Users\\cuong.nguyen\\Desktop\\Exec\\AKBC.xml'
filexml['AKiBC']  = 'C:\\Users\\cuong.nguyen\\Desktop\\Exec\\AKiBC.xml'
filexml['GBC']    = 'C:\\Users\\cuong.nguyen\\Desktop\\Exec\\GBC.xml'
filexml['HHBC']   = 'C:\\Users\\cuong.nguyen\\Desktop\\Exec\\HHBC.xml'
filexml['TBC']    = 'C:\\Users\\cuong.nguyen\\Desktop\\Exec\\TBC.xml'
const locationstore = []
locationstore['Success'] = 'C:\\Users\\cuong.nguyen\\GG\\StoreRequest\\Success\\'
locationstore['Fail'] = 'C:\\Users\\cuong.nguyen\\GG\\StoreRequest\\Fail\\'
const watcher = chokidar.watch('C:\\Users\\cuong.nguyen\\GG\\MasterData', {
    ignored: /(^|[\/\\])\../, // ignore dotfiles
    persistent: true
  });
  // Add event listeners.
 watcher.on('all' , (e,dirfile) => {
    const namefile = path.basename(dirfile).split(/\.txt/).shift()
    var extens = path.basename(dirfile).split('.').pop()
    if(e === 'add'){
      if(extens === 'txt'){
         var company = dirfile.split(/\\MasterData\\/).pop().split('\\').shift()       
         var email = namefile.split('-').pop()
         var cardcode = namefile.split('-').shift()
         writeFileXML(company,path.basename(dirfile),cardcode,email,dirfile)
      }
    }
 })
 function executeXML(cardcode,email,company,dirfile,namefile){
   process.exec('powershell -Command "DTW -s '+filexml[company]+'; sleep(6);$ws = New-Object -ComObject wscript.shell; $ws.SendKeys(\'{ENTER}\');$ws.SendKeys(\'{ENTER}\');$ws.SendKeys(\'{ENTER}\');sleep(3);$ws.SendKeys(\'{ENTER}\')"',(err, stdout , stderr) => {
      console.log("Đã thực thi ... Đợi check")
      checkResults(cardcode,email,dirfile,namefile)
   })
  }
function checkResults(cardcode,email,dirfile,namefile){
   var pool = new sql.ConnectionPool(config)
   let connected = pool.connect().then( pool => {
      return pool
   })
   .then( (pool) => {
     return pool.request()
      .input("cardcode",sql.VarChar,cardcode)   
      .query("select CardCode,CardName FROM [192.168.3.9].DATA_HBC_Test.dbo.OCRD a where a.CardCode = @cardcode")
      .then(result => { 
         // nếu >= 1 => success or somethingelse
         if(result.recordset.length >= 1){
            console.log('vào rồi')
            moveFileStore(dirfile,namefile,'Success')
            // sendResultCheck(email,'Thành công')
         }
         // Nếu not success i dont know =))
         else{
            console.log(dirfile)
            moveFileStore(dirfile,namefile,'Fail')
            // sendResultCheck(email,'Lỗi xảy ra')
         }
      })  
   })
}

function sendResultCheck(email,mess){
 const transporter = nodemail.createTransport({
   service: 'gmail',
   auth: {
     user: 'it@hbc.com.vn',
     pass: 'wbenbvzbyiofcuxa'
   }
 })
 const mailoption = {
   from: 'it@hbc.com.vn',
   to: email,
   subject: 'Thông báo thêm master data',
   text: mess
 }
 transporter.sendMail(mailoption,(error,info) => {
   if (error) {
      console.log(error);
    } else {
      console.log('Email sent: ' + info.response);
    }
 })
}

function writeFileXML(company,namefile,cardcode,email,dirfile){
   var content = '<Files><BusinessPartners>C:\\Users\\cuong.nguyen\\GG\\MasterData\\'+company+'\\'+namefile+'</BusinessPartners></Files>'
   var regex = /<Files><BusinessPartners>*.*<\/BusinessPartners><\/Files>/
   fs.readFile(filexml[company],(err,data) => {
      var stringreplace = data.toString().replace(regex,content)
      fs.writeFile(filexml[company],stringreplace,(err) => {
         if(err) console.log(err)
         else executeXML(cardcode,email,company,dirfile,namefile)
      })
   })
} 
 
function moveFileStore(dirfile,namefile,option){
   var cmd = 'move '+dirfile +' '+locationstore[option]+namefile
   console.log(cmd)
   process.exec(cmd)
}