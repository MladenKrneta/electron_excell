const readXlsxFile = require('read-excel-file/node');
const { Client } = require("pg");
const moment = require('moment');
const aws = require('./config.json')
const { dialog } = remote;

const schema = {
    'NADNEVAK': {
      prop: 'nadnevak',
      type: Date
    },
    'BRTEMELJNICE': {
      prop: 'brtemeljnice',
      type: String,
    },
    'OPIS_ISPRAVA': {
        prop: 'opis_isprava',
        type: String,
    },
    'U_GOTOVINI': {
        prop: 'u_gotovini',
        type: Number,
    },
    'ZIRO_RACUN': {
        prop: 'ziro_racun',
        type: Number,
    },
    'U_NARAVI': {
        prop: 'u_naravi',
        type: Number,
    },
    'PDV': {
        prop: 'pdv',
        type: Number,
    },
    'UKUPNO': {
        prop: 'ukupno',
        type: Number,
    },
    'U_GOTOVINI1': {
        prop: 'u_gotovini1',
        type: Number,
    },
    'ZIRO_RACUN1': {
        prop: 'ziro_racun1',
        type: Number,
    },
    'U_NARAVI1': {
        prop: 'u_naravi1',
        type: Number,
    },
    'PDV1': {
        prop: 'pdv1',
        type: Number,
    },
    'ST1T1_I_5': {
        prop: 'st1t1_i_5',
        type: Number,
    },
    'DOP_IZD': {
        prop: 'dop_izd',
        type: Number,
    },
    'TRANID': {
        prop: 'tranid',
        type: String,
    },
    'DIN': {
        prop: 'din',
        type: Number,
    },
  }
  
  function ExcelDateToJSDate(date) {
    return new Date(Math.round((date - 25569)*86400*1000));
  }

const upload = async (excel, name) =>{
    try{
        
        const client = new Client(aws);
        await client.connect();
    
        readXlsxFile(excel, { schema }).then((rows) => {
    
            let t = "INSERT INTO obai_ransom.synesis_ura_ira(ime, nadnevak, brtemeljnice, opis_isprava, u_gotovini, ziro_racun, u_naravi, pdv, ukupno, u_gotovini1, ziro_racun1, u_naravi1, pdv1, st1t1_i_5, dop_izd, tranid, din) VALUES";
    
            let v = ''
            rows.rows.map((x,index)=>{
    
                v = v + "(" +
                "'" + name + "'" + "," +
                "'" + moment(x.nadnevak).format('DD-MM-YYYY') + "'" + "," +
                ( x.brtemeljnice? ("'"+x.brtemeljnice + "'" ) : ('null'))+"," +
                ( x.opis_isprava? ("'"+x.opis_isprava + "'" ) : ('null'))+"," +
                "'" + x.u_gotovini + "'" + "," +
                "'" + x.ziro_racun + "'" + "," +
                "'" + x.u_naravi + "'" + "," +
                "'" + x.pdv + "'" + "," +
                "'" + x.ukupno + "'" + "," +
                "'" + x.u_gotovini1 + "'" + "," +
                "'" + x.ziro_racun1 + "'" + "," +
                "'" + x.u_naravi1 + "'" + "," +
                "'" + x.pdv1 + "'" + "," +
                "'" + x.st1t1_i_5 + "'" + "," +
                "'" + x.dop_izd + "'" + "," +
                "'" + x.tranid + "'" + "," +
                "'" + x.din + "'"  + ")" 
                
                if(index + 1===rows.rows.length){
                    v = v+";" 
                }else{
                    v = v+","
                }
            })
            kveri = t + v
            console.log(kveri)
            client.query(kveri)
            .then((res)=>{
              dialog.showMessageBox({ message: "Instert uspjesno napravljen!\r\nResponse: "+res.command + "\r\nCount: " + res.rowCount})
    
            })
            .catch((e)=>{
              console.log( e)
              dialog.showMessageBox({ message: 'An Error has occured! Read console for full details or Contact the creator :P'})
            })
        
        })
    
        } 
        catch (err) {
            console.log(err);
            // ... error checks
        }
}


exports.upload = upload;