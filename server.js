var excelbuilder = require('msexcel-builder');

// Create a new workbook file in current working-path
var workbook = excelbuilder.createWorkbook('./', 'sample.xlsx')

// Create a new worksheet with 14 columns and 12 rows
var sheet1 = workbook.createSheet('sheet1', 14, 50);

// column width
sheet1.width(2, 50);
sheet1.width(3, 30);
sheet1.width(4, 30);
sheet1.width(5, 25);
sheet1.width(6, 30);



// Fill some data

// row
sheet1.set(1, 1, 'QCID');
sheet1.set(2, 1, 'Production QC Steps');
sheet1.set(3, 1, 'E-Discovery Instructions');
sheet1.set(4, 1, 'Format of QC Step Answer');
sheet1.set(5, 1, 'Vendor QC');
sheet1.set(6, 1, 'Vendor Sample Bates Numbers');

// color of row
sheet1.fill(1, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(2, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(3, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(4, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(5, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(6, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(7, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(8, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(9, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(10, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(11, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(12, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(13, 1, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(14, 1, {type:'solid',fgColor:'8',bgColor:'64'});


// row
sheet1.set(1, 2, '1');
sheet1.set(2, 2, '');
sheet1.set(5, 2, 'qcid1');
sheet1.set(6, 2, '');

// row
sheet1.set(1, 3, '2');
sheet1.set(2, 3, '');
sheet1.set(5, 3, 'qcid2');
sheet1.set(6, 3, '');

// row
sheet1.set(1, 4, '4');
sheet1.set(2, 4, '');
sheet1.set(5, 4, 'qcid4');
sheet1.set(6, 4, '');

// row
sheet1.set(1, 5, '5');
sheet1.set(2, 5, '');
sheet1.set(5, 5, 'qcid5');
sheet1.set(6, 5, '');

// row
sheet1.set(1, 5, '5');
sheet1.set(2, 5, '');
sheet1.set(5, 5, 'qcid5');
sheet1.set(6, 5, 'qcid5Bates');

// row
sheet1.set(1, 6, '7');
sheet1.set(2, 6, '');
sheet1.set(5, 6, 'qcid7');
sheet1.set(6, 6, '');

// row
sheet1.set(1, 7, '60');
sheet1.set(2, 7, '');
sheet1.set(5, 7, 'qcid60');
sheet1.set(6, 7, '');

// row
sheet1.set(1, 8, '8');
sheet1.set(2, 8, '');
sheet1.set(5, 8, 'qcid8Count qcid8');
sheet1.set(6, 8, '');

// row
sheet1.set(1, 9, '9');
sheet1.set(2, 9, '');
sheet1.set(5, 9, 'qcid9Count qcid9');
sheet1.set(6, 9, '');

// row
sheet1.set(1, 10, '10');
sheet1.set(2, 10, '');
sheet1.set(5, 10, 'qcid10Count qcid10');
sheet1.set(6, 10, '');

// row
sheet1.set(1, 11, '11');
sheet1.set(2, 11, '');
sheet1.set(5, 11, 'qcid11Count qcid11');
sheet1.set(6, 11, '');



// row
sheet1.set(2, 12, 'DAT, OPT, and METADATA');

// color of row
sheet1.fill(1, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(2, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(3, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(4, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(5, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(6, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(7, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(8, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(9, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(10, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(11, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(12, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(13, 12, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(14, 12, {type:'solid',fgColor:'8',bgColor:'64'});


// row
sheet1.set(1, 13, '15');
sheet1.set(2, 13, '');
sheet1.set(5, 13, 'qcid15');
sheet1.set(6, 13, '');

// row
sheet1.set(1, 14, '16');
sheet1.set(2, 14, '');
sheet1.set(5, 14, 'qcid16');
sheet1.set(6, 14, '');

// row
sheet1.set(1, 15, '17');
sheet1.set(2, 15, '');
sheet1.set(5, 15, 'qcid17');
sheet1.set(6, 15, '');

// row
sheet1.set(1, 16, '20');
sheet1.set(2, 16, '');
sheet1.set(5, 16, 'qcid20');
sheet1.set(6, 16, '');

// row
sheet1.set(1, 17, '22');
sheet1.set(2, 17, '');
sheet1.set(5, 17, 'qcid22');
sheet1.set(6, 17, '');

// row
sheet1.set(1, 18, '25');
sheet1.set(2, 18, '');
sheet1.set(5, 18, 'qcid25');
sheet1.set(6, 18, '');

// row
sheet1.set(1, 19, '26');
sheet1.set(2, 19, '');
sheet1.set(5, 19, 'qcid26');
sheet1.set(6, 19, '');

// row
sheet1.set(1, 20, '27');
sheet1.set(2, 20, '');
sheet1.set(5, 20, 'qcid27');
sheet1.set(6, 20, '');




// row
sheet1.set(2, 21, 'REDACTIONS');

// color of row
sheet1.fill(1, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(2, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(3, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(4, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(5, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(6, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(7, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(8, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(9, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(10, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(11, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(12, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(13, 21, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(14, 21, {type:'solid',fgColor:'8',bgColor:'64'});


// row
sheet1.set(1, 22, '28');
sheet1.set(2, 22, '');
sheet1.set(5, 22, 'qcid28');
sheet1.set(6, 22, '');

// row
sheet1.set(1, 23, '29');
sheet1.set(2, 23, '');
sheet1.set(5, 23, 'qcid29');
sheet1.set(6, 23, '');

// row
sheet1.set(1, 24, '30');
sheet1.set(2, 24, '');
sheet1.set(5, 24, 'qcid30');
sheet1.set(6, 24, '');

// row
sheet1.set(1, 25, '31');
sheet1.set(2, 25, '');
sheet1.set(5, 25, 'qcid31');
sheet1.set(6, 25, '');

// row
sheet1.set(1, 26, '32');
sheet1.set(2, 26, '');
sheet1.set(5, 26, 'qcid32');
sheet1.set(6, 26, '');




// row
sheet1.set(2, 27, 'TEXT FILES');

// color of row
sheet1.fill(1, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(2, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(3, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(4, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(5, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(6, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(7, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(8, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(9, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(10, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(11, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(12, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(13, 27, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(14, 27, {type:'solid',fgColor:'8',bgColor:'64'});


// row
sheet1.set(1, 28, '33');
sheet1.set(2, 28, '');
sheet1.set(5, 28, 'qcid33');
sheet1.set(6, 28, 'qcid33Bates');

// row
sheet1.set(1, 29, '34');
sheet1.set(2, 29, '');
sheet1.set(5, 29, 'qcid34');
sheet1.set(6, 29, 'qcid34Bates');

// row
sheet1.set(1, 30, '35');
sheet1.set(2, 30, '');
sheet1.set(5, 30, 'qcid35');
sheet1.set(6, 30, 'qcid35Bates');

// row
sheet1.set(1, 31, '36');
sheet1.set(2, 31, '');
sheet1.set(5, 31, 'qcid36');
sheet1.set(6, 31, 'qcid36Bates');



// row
sheet1.set(2, 32, 'IMAGES');

// color of row
sheet1.fill(1, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(2, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(3, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(4, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(5, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(6, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(7, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(8, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(9, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(10, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(11, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(12, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(13, 32, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(14, 32, {type:'solid',fgColor:'8',bgColor:'64'});


// row
sheet1.set(1, 33, '37');
sheet1.set(2, 33, '');
sheet1.set(5, 33, 'qcid37');
sheet1.set(6, 33, 'qcid37Bates');

// row
sheet1.set(1, 34, '38');
sheet1.set(2, 34, '');
sheet1.set(5, 34, 'qcid38');
sheet1.set(6, 34, 'qcid38Bates');

// row
sheet1.set(1, 35, '39');
sheet1.set(2, 35, '');
sheet1.set(5, 35, 'qcid39');
sheet1.set(6, 35, '');

// row
sheet1.set(1, 36, '40');
sheet1.set(2, 36, '');
sheet1.set(5, 36, 'qcid40');
sheet1.set(6, 36, '');

// row
sheet1.set(1, 37, '42');
sheet1.set(2, 37, '');
sheet1.set(5, 37, 'qcid42');
sheet1.set(6, 37, 'qcid42Bates');

// row
sheet1.set(1, 38, '44');
sheet1.set(2, 38, '');
sheet1.set(5, 38, 'qcid44');
sheet1.set(6, 38, 'qcid44Bates');

// row
sheet1.set(1, 39, '45');
sheet1.set(2, 39, '');
sheet1.set(5, 39, 'qcid45');
sheet1.set(6, 39, 'qcid45Bates');




// row
sheet1.set(2, 40, 'NATIVE FILES');

// color of row
sheet1.fill(1, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(2, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(3, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(4, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(5, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(6, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(7, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(8, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(9, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(10, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(11, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(12, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(13, 40, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(14, 40, {type:'solid',fgColor:'8',bgColor:'64'});

// row
sheet1.set(1, 41, '47');
sheet1.set(2, 41, '');
sheet1.set(5, 41, 'qcid47');
sheet1.set(6, 41, '');

// row
sheet1.set(1, 42, '48');
sheet1.set(2, 42, '');
sheet1.set(5, 42, 'qcid48');
sheet1.set(6, 42, 'qcid48Bates');




// row
sheet1.set(2, 43, 'PRODUCTION LOG');

// color of row
sheet1.fill(1, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(2, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(3, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(4, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(5, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(6, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(7, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(8, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(9, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(10, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(11, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(12, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(13, 43, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(14, 43, {type:'solid',fgColor:'8',bgColor:'64'});


// row
sheet1.set(1, 44, '53');
sheet1.set(2, 44, '');
sheet1.set(5, 44, 'qcid53');
sheet1.set(6, 44, '');




// row
sheet1.set(2, 45, 'MDL (For Production)');

// color of row
sheet1.fill(1, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(2, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(3, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(4, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(5, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(6, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(7, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(8, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(9, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(10, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(11, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(12, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(13, 45, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(14, 45, {type:'solid',fgColor:'8',bgColor:'64'});


// row
sheet1.set(1, 46, '53A');
sheet1.set(2, 46, '');
sheet1.set(5, 46, 'qcid53A');
sheet1.set(6, 46, '');



// row
sheet1.set(2, 47, 'SECONDARY DAT (For Production)');

// color of row
sheet1.fill(1, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(2, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(3, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(4, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(5, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(6, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(7, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(8, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(9, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(10, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(11, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(12, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(13, 47, {type:'solid',fgColor:'8',bgColor:'64'});
sheet1.fill(14, 47, {type:'solid',fgColor:'8',bgColor:'64'});


// row
sheet1.set(1, 48, '54');
sheet1.set(2, 48, '');
sheet1.set(5, 48, 'qcid54Count qcid54');
sheet1.set(6, 48, '');

// row
sheet1.set(1, 49, '55');
sheet1.set(2, 49, '');
sheet1.set(5, 49, 'qcid55');
sheet1.set(6, 49, '');

// row
sheet1.set(1, 50, '56');
sheet1.set(2, 50, '');
sheet1.set(5, 50, 'qcid56');
sheet1.set(6, 50, '');




// Save it
workbook.save(function(ok){
    if (!ok)
        workbook.cancel();
    else
        console.log('congratulations, your workbook created');
});
