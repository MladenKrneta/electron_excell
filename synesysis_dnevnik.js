const readXlsxFile = require('read-excel-file/node');
const { Client } = require("pg");
const moment = require('moment');
const aws = require('./config.json')
const { dialog } = remote;

const schema = {
    "DATUM": {
      prop: 'datum',
      type: Date
    },
    "PRIORITY": {
      prop: 'priority',
      type: Number,
    },
    "DIN": {
        prop: 'din',
        type: Number,
    },
    "DOKUMENT": {
        prop: 'dokument',
        type: String,
    },
    "BROJ": {
        prop: 'broj',
        type: Number,
    },
    "ORGANIZACIJSKI_DIO": {
        prop: 'organizacijski_dio',
        type: String,
    },
    "KONTO": {
        prop: 'konto',
        type: String,
    },
    "SIFRA": {
        prop: 'sifra',
        type: Number,
    },
    "NAZIV": {
        prop: 'naziv',
        type: String,
    },
    "OPIS_KNJIZENJA": {
        prop: 'opis_knjizenja',
        type: String,
    },
    "DUGUJE": {
        prop: 'duguje',
        type: Number,
    },
    "POTRAZUJE": {
        prop: 'potrazuje',
        type: Number,
    },
    "KONTROLOR": {
        prop: 'kontrolor',
        type: String,
    }
  }
  
  function ExcelDateToJSDate(date) {
    return new Date(Math.round((date - 25569)*86400*1000));
  }

const upload = async (excel, name) =>{
    try{
        
        const client = new Client(aws);
        await client.connect();
    
        readXlsxFile(excel, { schema }).then((rows) => {
            let t = "INSERT INTO obai_ransom.synesis_dnevnici(ime, datum, priority, din, dokument, broj, organizacijski_dio, konto, sifra, naziv, opis_knjizenja, duguje, potrazuje, kontrolor) VALUES";
    
            let v = ''
            rows.rows.map((x,index)=>{
                v = v + "(" +
                "'" + name + "'" + "," +
                "'" + moment(x.datum).format('DD-MM-YYYY') + "'" + "," +
                "'" + x.priority + "'" + "," +
                "'" + x.din + "'" + "," +
                "'" + x.dokument + "'" +  "," +
                "'" + x.broj + "'" +  "," +
                ( x.organizacijski_dio? ("'"+x.organizacijski_dio + "'" ) : ('null'))+"," +
                "'" + x.konto + "'" +  "," +
                "'" + x.sifra + "'" +  "," +
                ( x.naziv? ("'"+x.naziv + "'" ) : ('null'))+"," +
                ( x.opis_knjizenja? ("'"+x.opis_knjizenja + "'" ) : ('null'))+"," +
                "'" + x.duguje + "'" +  "," +
                "'" + x.potrazuje + "'" +  "," +
                "'" + x.kontrolor + "'" + ")" 
                
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