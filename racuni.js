const readXlsxFile = require('read-excel-file/node');
const { Client } = require("pg");
const moment = require('moment');
const aws = require('./config.json')
const { dialog } = remote;

const schema = {
    'Lokal': {
      prop: 'lokal',
      type: String
    },
    'Blagajna': {
      prop: 'blagajna',
      type: String,
    },
    'Knjigovodstveni datum': {
        prop: 'knjigovodstveni_datum',
        type: Date,
    },
    'Datum i vrijeme': {
        prop: 'stamp',
        type: String,
    },
    'Način plaćanja': {
        prop: 'nacinplacanja',
        type: String,
    },
    'Fiskalni broj računa': {
        prop: 'fiskalnibroj',
        type: String,
    },
    'Izdao': {
        prop: 'izdao',
        type: String,
    },
    'Kupac': {
        prop: 'kupac',
        type: String,
    },
    'Porezni broj kupca': {
        prop: 'poreznibrojkupca',
        type: String,
    },
    'PDV': {
        prop: 'pdv',
        type: Number,
    },
    'PNP': {
        prop: 'pnp',
        type: Number,
    },
    'UkupnoRacun': {
        prop: 'ukupnoracun',
        type: Number,
    },
    'Oznaka porezne grupe': {
        prop: 'oznakaporeznegrupe',
        type: String,
    },
    'Šifra': {
        prop: 'sifra',
        type: String,
    },
    'Artikl': {
        prop: 'artikl',
        type: String,
    },
    'Količina': {
        prop: 'kolicina',
        type: Number,
    },
    'Cijena': {
        prop: 'cijena',
        type: Number,
    },
    'Cijena sa popustom': {
        prop: 'cijenapopust',
        type: Number,
    },
    'Ukupno popusta': {
        prop: 'ukupnopopust',
        type: Number,
    },
    'Ukupno neto': {
        prop: 'ukupnoneto',
        type: Number,
    },
    'Ukupno': {
        prop: 'ukupno',
        type: Number,
    },
  }
  
  function ExcelDateToJSDate(date) {
    return new Date(Math.round((date - 25569)*86400*1000));
  }

const upload = async (excel) =>{
    try{
        
        const client = new Client(aws);
        await client.connect();
    
        readXlsxFile(excel, { schema }).then((rows) => {
    
            t = "INSERT INTO sumarak.remaris_izlazni VALUES";
    
            let v = ''
            rows.rows.map((x,index)=>{
    
                v = v + "(" +
                "'" + x.lokal + "'" + "," +
                "'" + x.blagajna + "'" + "," +
                "'" + moment(ExcelDateToJSDate(x.stamp)).format() + "'" + "," +
                "'" + x.nacinplacanja + "'" + "," +
                "'" + x.fiskalnibroj + "'" +  "," +
                "'" + x.izdao + "'" +  "," +
                ( x.kupac? ("'"+x.kupac + "'" ) : ('null'))+"," +
                ( x.poreznibrojkupca? ("'"+x.poreznibrojkupca + "'" ) : ('null'))+"," +
                "'" + x.pdv + "'" +  "," +
                "'" + x.pnp + "'" +  "," +
                "'" + x.ukupnoracun + "'" +  "," +
                "'" + x.oznakaporeznegrupe + "'" +  "," +
                "'" + x.sifra + "'" +  "," +
                "'" + x.artikl + "'" +  "," +
                "'" + x.kolicina + "'" +  "," +
                "'" + x.cijena + "'" +  "," +
                "'" + x.cijenapopust + "'" +  "," +
                "'" + x.ukupnopopust + "'" +  "," +
                "'" + x.ukupnoneto + "'" +  "," +
                "'" + x.ukupno + "'" +  ")" 
                
                if(index + 1===rows.rows.length){
                    v = v+";" 
                }else{
                    v = v+","
                }
            })
            kveri = t + v
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