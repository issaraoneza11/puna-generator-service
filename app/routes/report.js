const express = require('express');


const Responsedata = require('../middleware/response');

const databaseContextPg = require("database-context-pg");
const connectionSetting = require("../dbconnect");

const connectionConfig = connectionSetting.config;
const condb = new databaseContextPg(connectionConfig);
const { v4: uuidv4 } = require("uuid");
const router = express.Router();
const moment = require('moment');
var multiparty = require('multiparty');
const fs = require("fs");
const { Get_font_pdf_th2 } = require("../font/pdf_font");
const cordiaNormal = require("../font/CordiaNew-normal.js");
const cordiaBold = require("../font/CordiaNew-bold.js");
const tahomaNormal = require("../font/Tahoma-normal.js");
const { createPdfTools } = require("../lib/reportGenDocPdf");
const { log } = require('console');


/* const formVat10 = require("../form_report/formVat10.jpg");
 */

router.get('/', function (req, res, next) {
    console.log('เข้าเเล้ว');
    /*  res.render('index', {
         title: 'TTT GAME SERVICE DEMO'
     }); */
});

const calPage = (count, size) => {
    return Math.ceil(count / size);
};

const calSkip = (page, size) => {
    return (page - 1) * size;
};


router.get('/test', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {



        let tempRes = {
            count: 1,
            data: 1
        }

        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                /*   { name: 'THSarabunNew', vfsName: 'THSarabunNew.ttf', data: Get_font_pdf_th2(), style: 'bold' }, */
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },
                { name: 'tahoma', vfsName: 'tahoma Normal.ttf', data: stripDataUrl(tahomaNormal), style: 'normal' },


            ]
        });
        const { Shot, CutPadding, state, splitText } = pdf;
        Shot('sf', 'CordiaNew', 'bold', 18);

        Shot("t", "test ว่าได้ไหม", 10, 50, 'r')
        state.y = 10
        Shot("t", "test ว่าได้ไหม", CutPadding('w', 50), state.y, 'c');
        state.y += 6
        Shot('sf', 'CordiaNew', 'normal', 12);
        /* 
         Shot("t", "test มันซ้ายไหม", CutPadding('w', 50), state.y, 'l');
         state.y += 6
         Shot("fs", 10);
         Shot("t", "test ขวาสุดสิ", CutPadding('w', 100), state.y, 'r');
         state.y += 6
         Shot("tc", 255, 0, 0);
         Shot("t", "test สีสิ", CutPadding('w', 100), state.y, 'l');
         state.y += 6
         Shot('sf', 'CordiaNew', 'bold', 18);
         Shot("t", "a b c A B C", CutPadding('w', 50), state.y, 'l');
 
         state.y += 6
         Shot('sf', 'CordiaNew', 'bold', 18);
         Shot("t", "สพานสูงที่ลิมน้ำ", CutPadding('w', 50), state.y, 'l');
         let cutText = splitText(`ลองเช็คว่าตัดได้จริงไหม แบบ ยาวๆๆๆๆๆๆๆๆมันได้ไหกกกกกกกกกกกกกกกกกกกกกกกกกกกกหหหหหหหหหหม ยาวๆๆๆๆๆๆๆๆมันได้ไหกกกกกกกกกกกกกกกกกกกกกกกกกกกกหหหหหหหหหหม ยาวๆๆๆๆๆๆๆๆมันได้ไหกกกกกกกกกกกกกกกกกกกกกกกกกกกกหหหหหหหหหหม ยาวๆๆๆๆๆๆๆๆมันได้ไหกกกกกกกกกกกกกกกกกกกกกกกกกกกกหหหหหหหหหหม ยาวๆๆๆๆๆๆๆๆมันได้ไหกกกกกกกกกกกกกกกกกกกกกกกกกกกกหหหหหหหหหหม`, CutPadding('w', 20, 10), 2);
         console.log(cutText);
 
         state.y += 20
         Shot('sf', 'CordiaNew', 'normal', 35);
         Shot("t", "สภานที่สูงนะบึงการลองหม้าย a b c A B C", CutPadding('w', 20), state.y, 'l');
         state.y += 20
         Shot('sf', 'CordiaNew', 'bold', 35);
 <<<<<<< HEAD
         Shot("t", "สภานที่สูงนะบึงการลองหม้าย a b c A B C", CutPadding('w', 20), state.y, 'l');
         state.y += 20
         Shot('sf', 'tahoma', 'normal', 35);
         Shot("t", "สภานที่สูงนะบึงการลองหม้าย a b c A B C", CutPadding('w', 20), state.y, 'l');
 =======
         Shot("t", "สภานที่สูงนะบึงการลองหม้าย a b c A B C", CutPadding('w', 20), state.y, 'l'); */

        const buffer = await Shot("save", 'ลองเปลี่ยนดู');
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});
function getBase64Image(path) {
    return new Promise((resolve, reject) => {
        fs.readFile(path, (err, data) => {
            if (err) reject(err);
            else resolve(`data:image/jpeg;base64,${data.toString("base64")}`);
        });
    });
}





router.post('/formVat10', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;




        // ใช้งาน:
        const form_background = await getBase64Image('./app/form_report/formVat10.jpg');


        // console.log('base64',base64);
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                /*    { name: 'THSarabunNew', vfsName: 'THSarabunNew.ttf', data: Get_font_pdf_th2(), style: 'bold' }, */
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },
            ]
        });
        const { Shot, CutPadding, state, resizeText } = pdf;
        Shot("i", form_background, 0, 0, 210, 297);
        Shot("i", model.qrcode_base64, 18, 40, 30, 30);//qr_code
        Shot("i", model.barcode_base64, 55, 42, 130, 18);//bar_code
        state.y = 19;
        Shot("fs", 16);
        /*         Shot('sf', 'CordiaNew', 'bold', 16); */
        Shot('sf', 'CordiaNew', 'bold', 16);
        Shot("tb", model.declaration_number || '', CutPadding('w', 50), state.y, 'c', 2);



        state.y = 87.5;
        let reSizetxt = resizeText(model.name || '', 175, 16);
        Shot("fs", reSizetxt);
        Shot("tb", model.name || '', CutPadding('w', 6), state.y, 'l', 1);
        Shot('sf', 'CordiaNew', 'bold', 16);
        state.y += 10

        Shot("tb", model.ref_1 || '', CutPadding('w', 33), state.y, 'l', 2);


        state.y += 10

        Shot("tb", model.ref_2 || '', CutPadding('w', 33), state.y, 'l', 2);


        state.y += 10

        Shot("tb", model.price || '', CutPadding('w', 20), state.y, 'l', 2);





        const filename = req?.query?.file_name || "แบบฟอร์มชำระค่าภาษี"; // ชื่อไฟล์เป็นภาษาไทย
        //  console.log('filename');
        const buffer = await Shot("save", filename);
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});

router.post('/importDeclaration', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;

        // ใช้งาน:
        const form_background1 = await getBase64Image('./app/form_report/importDecalration1.jpg');
        const form_background2 = await getBase64Image('./app/form_report/importDecalration2.jpg');
        const form_background3 = await getBase64Image('./app/form_report/importDecalration3.jpg');

        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },
            ]
        });

        const { Shot, CutPadding, state, splitText, resizeText, addPage, doc } = pdf;


        // หน้า 1 (ใบแรก)

        Shot('sf', 'CordiaNew', 'normal', 10);
        Shot("i", form_background1, 0, 0, 210, 297);
        Shot("i", model.barcode_base64, 157, 4, 35, 10); // barcode
        state.y = 16;
        Shot("fs", 10);

        Shot("t", model?.is_privilege ? 'ใช้สิทธิประโยชน์' : '', 16, state.y, 'l');

        state.y = 21;
        Shot("t", model?.document_control?.document_type || '', CutPadding('w', 56), state.y, 'l');// ประเภทใบขน
        state.y = 26;
        Shot("t", model?.document_control?.duty_exemption || '', CutPadding('w', 46.5), state.y, 'l');// ยกเว้นค่าภาษีอากร
        Shot("fs", 14);
        state.y = 22;
        Shot("t", model?.document_control?.declaration_number || '', CutPadding('w', 87), state.y, 'l'); //เลขที่ใบขน
        state.y = 27;
        Shot("t", model?.document_control?.reference_number || '', CutPadding('w', 80), state.y, 'l'); //เลขที่ใบขน

        //สั่งการตรวจ
        Shot("fs", 10);
        let str = '';
        let font_size = 10;
        let max_line = 3;

        const timeLine = model?.time_line_detail || [];
        const timeLineLen = timeLine.length || 0;
        let new_font = timeLineLen ? (max_line / timeLineLen) * font_size : font_size;
        let line_height = new_font / 3;
        state.y = 25;
        if (new_font > 10) {
            new_font = 10;
        }

        Shot("fs", new_font);

        for (let item of timeLine) {
            str += (item?.date_time || '') + ' ' + (item?.description || '') + '\n';
        }
        let cutText = splitText(str, CutPadding('w', 34));

        for (let item2 of cutText) {
            Shot("t", item2, 18, state.y, 'l');
            state.y += line_height;
        }

        // ตารางอากร
        state.y = 37.5;
        Shot("fs", 14);
        //อาการขาเข้า 
        Shot("t", model?.duty?.import?.duty_amount || '', CutPadding('w', 83), state.y, 'r'); // ค่าภาษีอากร
        Shot("t", model?.duty?.import?.deposit_amount || '', CutPadding('w', 102), state.y, 'r'); //เงินประกัน
        state.y += 8;
        //ภาษีสรรพสามิต 
        Shot("t", model?.duty?.excise_duty?.duty_amount || '', CutPadding('w', 83), state.y, 'r'); // ค่าภาษีอากร
        Shot("t", model?.duty?.excise_duty?.duty_amount || '', CutPadding('w', 102), state.y, 'r'); //เงินประกัน
        state.y += 8;
        //ภาษีเพื่อมหาดไทย 
        Shot("t", model?.duty?.interior_duty?.duty_amount || '', CutPadding('w', 83), state.y, 'r'); // ค่าภาษีอากร
        Shot("t", model?.duty?.interior_duty?.duty_amount || '', CutPadding('w', 102), state.y, 'r'); //เงินประกัน
        state.y += 8;
        //ภาษีมูลค่าเพิ่ม 
        Shot("t", model?.duty?.vat_duty?.duty_amount || '', CutPadding('w', 83), state.y, 'r'); // ค่าภาษีอากร
        Shot("t", model?.duty?.vat_duty?.duty_amount || '', CutPadding('w', 102), state.y, 'r'); //เงินประกัน
        state.y += 8;
        //ภาษีและค่าธรรมเนียมอื่นๆ 
        Shot("t", model?.duty?.other_duty?.duty_amount || '', CutPadding('w', 83), state.y, 'r'); // ค่าภาษีอากร
        Shot("t", model?.duty?.other_duty?.duty_amount || '', CutPadding('w', 102), state.y, 'r'); //เงินประกัน
        state.y += 8;
        //รวมทั้งสิ้น 
        Shot("t", model?.duty?.summary?.duty_amount || '', CutPadding('w', 83), state.y, 'r'); // ค่าภาษีอากร
        Shot("t", model?.duty?.summary?.duty_amount || '', CutPadding('w', 102), state.y, 'r'); //เงินประกัน

        Shot("fs", 10);
        let account_product = '';
        let index_account_product = 0;

        const accountList = model?.account_product_price || [];
        const accountLen = accountList.length;

        for (let item of accountList) {
            if (index_account_product + 1 == accountLen) {
                account_product += item;
            } else {
                account_product += item + ', ';
            }
            index_account_product++;
        }
        state.y = 86.5;

        let reSizetxt = resizeText(account_product, 110, 9);
        Shot("fs", reSizetxt);
        Shot("t", account_product || '', CutPadding('w', 74), state.y, 'c'); //เลขที่บัญชีราคาสินค้า

        state.y = 93;
        Shot("fs", 9);
        Shot("t", model?.document_control?.duty_exemption || '', CutPadding('w', 46.5), state.y, 'l');// ยกเว้นค่าภาษีอากร
        Shot("t", ('วิธีการชำระเงิน ' + (model?.document_control?.bank_info?.payment_method || '')) || '', CutPadding('w', 101.5), state.y, 'r');// วิธีการชำระเงิน
        state.y = 98.5;
        let TaxintenText = model?.document_control?.tax_incentive?.tax_incentive?.[0] || ''; // รหัสยกเว้นอากร
        let taxIncentiveFontSize = resizeText(TaxintenText, 78, 9);
        Shot("fs", taxIncentiveFontSize);
        Shot("t", TaxintenText, CutPadding('w', 46.5), state.y, 'l');

        state.y = 112;
        Shot("fs", 9);
        Shot("t", ('Type: ' + (model?.document_control?.bank_info?.guarantee_type || '')) || '', CutPadding('w', 46.5), state.y, 'l'); //ประเภทการชำระเงิน


        state.y = 121.5;
        Shot("t", model?.document_control?.origin_country_name || model?.document_control?.origin_country?.origin_country_name || '', CutPadding('w', 46.5), state.y, 'l');// ประเทศกำเนิด
        Shot("t", model?.document_control?.origin_country_code || model?.document_control?.origin_country?.origin_country_code || '', CutPadding('w', 72.5), state.y, 'r');// รหัสประเทศกำเนิด

        Shot("t", model?.document_control?.consignment_country_name || model?.document_control?.consignment_country?.consignment_country_name || '', CutPadding('w', 74.5), state.y, 'l');// ประเทศต้นทางบรรทุก
        Shot("t", model?.document_control?.consignment_country_code || model?.document_control?.consignment_country?.consignment_country_code || '', CutPadding('w', 101), state.y, 'r');// รหัสประเทศต้นทางบรรทุก

        state.y = 131.5;
        Shot("t", model?.document_control?.discharge_port_name || model?.document_control?.discharge_port?.discharge_port_name || '', CutPadding('w', 46), state.y, 'l');// ท่าเรือที่นำเข้า
        Shot("t", model?.document_control?.discharge_port_code || model?.document_control?.discharge_port?.discharge_port_code || '', CutPadding('w', 73), state.y, 'r');// รหัสท่าเรือที่นำเข้า

        Shot("t", model?.document_control?.release_port_name || model?.document_control?.release_port?.release_port_name || '', CutPadding('w', 74.5), state.y, 'l');// สถานที่ตรวจปล่อย
        Shot("t", model?.document_control?.release_port_code || model?.document_control?.release_port?.release_port_code || '', CutPadding('w', 101.5), state.y, 'r');// รหัสสถานที่ตรวจปล่อย

        state.y = 137.1;
        Shot("t", model?.document_control?.rate_of_exchange || model?.document_control?.currency_exchange?.exchang_rate + '' + model?.document_control?.currency_exchange?.currency, CutPadding('w', 83), state.y, 'l');// อัตราแลกเปลี่ยน


        Shot("fs", 9);
        state.y = 43;
        Shot("t", model?.document_control?.importer?.tax_number || '', CutPadding('w', 19.5), state.y, 'l');// เลขประจำตัวผู้เสียภาษีอากร
        Shot("t", model?.document_control?.importer?.branch || '', CutPadding('w', 44.5), state.y, 'r');// เลขประจำตัวผู้เสียภาษีอากร สาขา

        // ผู้นําของเข้า (ชื่อ ที่อยู่ โทรศัพท์)
        state.y = 47;
        Shot("t", model?.document_control?.importer?.english_name || '', 17, state.y, 'l');
        state.y += 3.5;
        Shot("t", model?.document_control?.importer?.thai_name || '', 17, state.y, 'l');
        state.y += 3.5;
        Shot("t", model?.document_control?.importer?.street_and_number || '', 17, state.y, 'l');
        state.y += 3.5;
        Shot("t", (model?.document_control?.importer?.sub_province || '') + ' ' + (model?.document_control?.importer?.district || ''), 17, state.y, 'l');
        state.y += 3.5;
        Shot("t", (model?.document_control?.importer?.provice || '') + ' ' + (model?.document_control?.importer?.postcode || ''), 17, state.y, 'l');

        //ชื่อและเลขที่บัตรผ่านพิธีการ
        state.y = 74;
        Shot("t", (model?.document_control?.authorised_person?.customsclearancename || '') + ' ' + (model?.document_control?.authorised_person?.customs_clearance_id_card || '') + ' / ' + (model?.document_control?.authorised_person?.customs_expire_date || ''), 18.5, state.y, 'l');

        //ตัวแทนออกของ
        state.y = 82;
        Shot("t", (model?.document_control?.agent?.tax_number || '') + (model?.document_control?.agent?.branch ? ' สาขา ' + (model?.document_control?.agent?.branch || '') : ''), 18.5, state.y, 'l');
        state.y += 4;
        Shot("t", (model?.document_control?.agent?.thai_name || ''), 18.5, state.y, 'l');

        //ใบตราส่งเลขที่
        state.y = 95;
        Shot("t", (model?.document_control?.bill_of_lading?.house || ''), 18.5, state.y, 'l');
        state.y += 3;
        Shot("t", (model?.document_control?.bill_of_lading?.master || ''), 18.5, state.y, 'l');

        //นำเข้าทาง
        state.y = 95;
        Shot("t", (model?.document_control?.border_transport_means?.mode_code || ''), CutPadding('w', 32.5), state.y, 'l');

        //ชื่อยานพาหนะ
        state.y = 105.5;
        Shot("t", (model?.document_control?.border_transport_means?.mode_name || model?.document_control?.border_transport_means?.vessel || ''), 18.5, state.y, 'l');
        // วันที่นำเข้า
        Shot("t", (model?.document_control?.border_transport_means?.arrival_date || ''), CutPadding('w', 37), state.y, 'c');

        //เครื่องหมายและเลขหมายหีบห่อ
        state.y = 114.5;
        let cutText_ = splitText((model?.document_control?.total_package?.shipping_marks || model?.document_control?.total_package?.shipping_marks || ''), 53);

        for (let item2 of cutText_) {
            if (item2) {
                Shot("t", item2, 18.5, state.y, 'l');
                state.y += 3;
            }
        }

        //จำนวนและลักษณะหีบห่อ
        state.y = 140;
        Shot("t", model.document_control.total_package.amount + ' ' + model.document_control.total_package.unit_code + ' ('
            + model.document_control.total_package.amount_text + ' ' + model.document_control.total_package.unit_code_text + ')', 20, state.y, 'l');
        /* Shot("t", (model?.document_control?.total_package?.amount || ''), CutPadding('w', 38.5), state.y, 'c'); */
        state.y = 120;
        Shot("t", (model?.document_control?.total_package?.unit_code_text || ''), CutPadding('w', 38.5), state.y, 'c');

        //จำนวนหีบห่อรวม (ตัวเลข) (ตัวอักษร)
        state.y = 141;
        Shot(
            (model?.document_control?.total_package?.amount || '') +
            ' ' +
            (model?.document_control?.total_package?.unit_code_text || '') +
            ' (' +
            (model?.document_control?.total_package?.amount_text || '') +
            ' ' +
            (model?.document_control?.total_package?.unit_code_text || '') +
            ')',
            18.5,
            state.y,
            'l'
        );

        // ====== Item List ======

        const pageKeys = Object.keys(model.detail || {}).filter(
            (k) => model.detail[k] && typeof model.detail[k] === 'object' && Array.isArray(model.detail[k].items)
        );

        // ====== ตั้งค่าจำนวนแผ่น (รวม invoice + ใบปิด) ======
        const hasInvoicePage = !!(model.detail?.page_invoice?.invoice && model.detail.page_invoice.invoice.length);
        const totalSheets = pageKeys.length + (hasInvoicePage ? 1 : 0) + 1; // +1 สำหรับใบปิด
        let sheetIndex = 1; // หน้าแรกคือแผ่นที่ 1

        let data_last_page = {};
        let indexPage = 0;

        for (let key of pageKeys) {

            // ------------------------
            // ใบแรก: ใส่ได้ 3 item
            // ------------------------
            if (indexPage == 0) {
                data_last_page = model?.detail[key];
                state.y = 162.5;

                for (let item of (model?.detail[key]?.items || [])) {
                    // แถวบนสุด (พิกัด, ราคา ตปท, อัตราภาษี ฯลฯ)
                    Shot("t", item?.item_number || '', 18.5, state.y, 'c'); //รายการที่
                    Shot("t", item?.tariffifo?.code || '', CutPadding('w', 1.2), state.y - 13, 'l'); //ประเภทพิกัด
                    Shot("t", item?.cif_Value?.foreign || item?.price_foreign?.price_foreign || '', CutPadding('w', 30), state.y - 13, 'r'); //ราคาของ (เงินต่างประเทศ)
                    Shot("t", item?.duty?.duty_rate || '', CutPadding('w', 45), state.y - 13, 'r'); //อัตราอากรขาเข้า
                    Shot("t", item?.duty?.duty_fee || '', CutPadding('w', 59), state.y - 13, 'r'); // ค่าธรรมเนียม
                    Shot("t", item?.duty?.excise_code || '', CutPadding('w', 68.5), state.y - 13, 'c'); // รหัสสินค้าสรรพสามิต
                    Shot("t", item?.duty?.duty_excise || '', CutPadding('w', 86.5), state.y - 13, 'r'); // ภาษีสรรพสามิต
                    Shot("t", item?.duty?.duty_base_vat || '', CutPadding('w', 101.5), state.y - 13, 'r'); // ฐานภาษีมูลค่าเพิ่ม

                    // แถวกลาง (รหัสสถิติ, ราคาบาท, ภาษีอื่นๆ ฯลฯ)
                    Shot("t", (item?.tariffifo?.sequence || item?.tariffifo?.statitic || '') + ' / ' + (item?.tariffifo?.sequence_unit || item?.tariffifo?.statitic_unit || ''), CutPadding('w', 3.9), state.y - 3.5, 'c'); //รหัสสถิติ หน่วย
                    Shot("t", (item?.cif_Value?.baht || item?.price_baht || ''), CutPadding('w', 30), state.y - 3.5, 'r'); //ราคาของ (บาท)
                    Shot("t", (item?.duty?.duty_amount || ''), CutPadding('w', 45.2), state.y - 3.5, 'r'); //อากรขาเข้าที่ชำระ
                    Shot("t", (item?.duty?.duty_other || ''), CutPadding('w', 59), state.y - 3.5, 'r'); //ภาษีอื่นๆ
                    Shot("t", (item?.duty?.excise_rate || ''), CutPadding('w', 72), state.y - 3.5, 'r'); //อัตราภาษีสรรพสามิต
                    Shot("t", (item?.duty?.duty_interior || ''), CutPadding('w', 87.5), state.y - 3.5, 'r'); //ภาษีเพื่อมหาดไทย
                    Shot("t", (item?.duty?.duty_vat || ''), CutPadding('w', 101.5), state.y - 3.5, 'r'); //ภาษีมูลค่าเพิ่ม

                    // แถวล่าง (สิทธิพิเศษ, น้ำหนัก, ปริมาณ, คำอธิบาย)
                    Shot("t", (item?.privilege_code || ''), CutPadding('w', 2.3), state.y + 6, 'c'); //รหัสสิทธิพิเศษ
                    Shot("t", (item?.productifo?.net_weight?.weight || '') + ' ' + (item?.productifo?.net_weight?.unit_code || ''), CutPadding('w', 30), state.y + 6, 'r'); //น้ำหนักสุทธิ
                    Shot("t", (item?.quantity?.quantity || '') + ' ' + (item?.quantity?.UnitCode || ''), CutPadding('w', 45.1), state.y + 6, 'r'); //ปริมาณ

                    //ชนิดของ
                    let reSizetxt_brand = resizeText((item?.productifo?.brandname || item?.productifo?.brand?.brand_name || ''), 50, 9);
                    Shot("fs", reSizetxt_brand);
                    Shot("t", (item?.productifo?.brandname || item?.productifo?.brand?.brand_name || ''), CutPadding('w', 51.5), state.y + 2.5, 'l');

                    //  รายละเอียดสินค้า: อังกฤษ / ไทย / ลักษณะ 1 / ลักษณะ 2 
                    const engName = item?.goods_description?.english || item?.productifo?.english || '';
                    const thaiName = item?.goods_description?.thai || item?.productifo?.thai || '';
                    const attr1 = item?.productifo?.product_attribute1 || '';
                    const attr2 = item?.productifo?.product_attribute2 || '';
                    const descX = CutPadding('w', 46);
                    const baseY = state.y + 4.5;
                    // ชื่ออังกฤษ
                    const fsEng = resizeText(engName, 92, 8);
                    Shot("fs", fsEng);
                    Shot("t", engName, descX, baseY, 'l');

                    // ชื่อไทย
                    const fsThai = resizeText(thaiName, 120, 8);
                    Shot("fs", fsThai);
                    Shot("t", thaiName, descX, baseY + 2, 'l');

                    // ลักษณะที่ 1
                    const fsAttr1 = resizeText(attr1, 90, 8);
                    Shot("fs", fsAttr1);
                    Shot("t", attr1, descX, baseY + 4, 'l');

                    // ลักษณะที่ 2
                    const fsAttr2 = resizeText(attr2, 90, 8);
                    Shot("fs", fsAttr2);
                    Shot("t", attr2, descX, baseY + 6, 'l');


                    Shot("fs", 9);
                    Shot("t", (item?.origin_country_code || item?.productifo?.origin_country_code || ''), CutPadding('w', 102), state.y + 2.5, 'r');
                    Shot("t", (item?.import_tax_incentives_name || item?.productifo?.import_tax_incentives_name || ''), CutPadding('w', 102), state.y + 6.5, 'r');
                    Shot("t", (item?.import_tax_incentives_id || item?.productifo?.import_tax_incentives_id || ''), CutPadding('w', 102), state.y + 10.5, 'r');

                    Shot("fs", 8);
                    //หมายเหตุ
                    const remarkNoCV = ((item?.nature_of_transacton || item?.productifo?.nature_of_transacton) + ' ' + (item?.remark || item?.productifo?.remark || '')).trim();
                    const remarkFontSize = resizeText(remarkNoCV, 80, 10);
                    Shot("fs", remarkFontSize);
                    Shot("t", remarkNoCV, CutPadding('w', 51.5), state.y + 15.5, 'l');

                    //ใบอนุญาตินำเข้าหรือหนังสือรับรอง
                    Shot("t", (item?.license_number || (item?.permit && item?.permit[0]?.permit_no) || ''), CutPadding('w', 2), state.y + 14.5, 'l');

                    let index_various = 0;
                    for (let item2 of (item?.various_values || [])) {
                        Shot("t", (item2 || ''), CutPadding('w', 44.5), state.y + 11 + (2.5 * index_various), 'r');
                        index_various++;
                    }

                    state.y += 36; //อิงจาก เลขในรายการที่
                }

                //summary (หน้าแรก)
                state.y = 262;

                Shot("t", model?.detail?.[key]?.footer?.incoterm || '', CutPadding('w', 5.5), state.y, 'l');

                state.y = 253.5;
                const pf = model?.detail[key]?.summary?.price_foreign;
                const pfText = pf ? `${pf.price_foreign_amount || ''} ${pf.price_foreign_unit || ''}`.trim() : '';
                Shot("t", pfText, CutPadding('w', 30), state.y, 'r');

                state.y = 264;
                Shot("t", (model?.detail[key]?.summary?.price_amount_baht || ''), CutPadding('w', 30), state.y, 'r'); //รวม ราคาของ (บาท)
                state.y = 259;
                Shot("fs", 15);
                Shot("t", (model?.detail[key]?.summary?.duty_amount || ''), CutPadding('w', 44.5), state.y, 'r'); // รวม ค่าอากรขาเข้า
                state.y = 262;
                Shot("fs", 10);
                const qty = model?.detail?.[key]?.summary?.total_quality;
                const qtyText = qty ? `${qty.quantity_value || ''} ${qty.quality_unit || ''}`.trim() : '';
                Shot("t", qtyText, CutPadding('w', 44.5), state.y, 'r');

                state.y = 256;
                Shot("fs", 14);
                Shot("t", (model?.detail[key]?.summary?.duty_fee || ''), CutPadding('w', 59), state.y, 'r'); // รวม ค่าธรรมเนียม
                state.y = 264.5;
                Shot("t", (model?.detail[key]?.summary?.other_duty || ''), CutPadding('w', 59), state.y, 'r'); // รวม ภาษีอื่นๆ

                state.y = 256;
                Shot("t", (model?.detail[key]?.summary?.excise_duty || ''), CutPadding('w', 87), state.y, 'r'); // รวม สรรพสามิต
                state.y = 264.5;
                Shot("t", (model?.detail[key]?.summary?.interior_duty || ''), CutPadding('w', 87), state.y, 'r'); // รวม ภาษีมหาดไทย
                state.y = 258;
                Shot("t", (model?.detail[key]?.summary?.vat_duty || ''), CutPadding('w', 102), state.y, 'r'); // รวม ภาษีมูลค่าเพิ่ม

                state.y = 271;
                Shot("t", (model?.detail[key]?.summary?.total_duty_amount || ''), CutPadding('w', 102), state.y, 'r'); // รวมค่าภาษีอากรทั้งสิ้น

                //footer หน้าแรก
                Shot("fs", 10);
                state.y = 272;
                Shot("t", (model?.detail[key]?.footer?.net_weight || model?.detail[key]?.footer?.total_net_weight?.net_weight_value || ''), CutPadding('w', 24.5), state.y, 'r'); // น้ำหนักสุทธิ
                state.y = 277;
                Shot("t", (model?.detail[key]?.footer?.gross_weight || model?.detail[key]?.footer?.total_gross_weight?.gross_weight_value || ''), CutPadding('w', 22.5), state.y, 'r'); // G.W.

                state.y = 269;
                for (let item of (model?.detail[key]?.footer?.total_various_values || [])) {
                    Shot("t", (item || ''), CutPadding('w', 26), state.y, 'l');
                    state.y += 3;
                }

                state.y = 282;
                Shot("t", (model?.detail[key]?.footer?.status || ''), CutPadding('w', -5), state.y, 'l'); // สถานะล่าสุด
                state.y = 286;
                Shot("t", (model?.detail[key]?.footer?.status_date || ''), CutPadding('w', -5), state.y, 'l'); // วันที่สถานะล่าสุด

                state.y = 287;
                Shot("t", (model?.detail[key]?.footer?.signature_name || model?.detail[key]?.footer?.signature?.signature_name || ''), CutPadding('w', 64), state.y, 'l'); // ลายมือชื่ออิเล็กทรอนิกส์
                state.y = 291;
                Shot("t", (model?.detail[key]?.footer?.signature_date || ''), CutPadding('w', 64), state.y, 'l'); // วันที่ยื่น

                // จบหน้าแรก -> แผ่นต่อไปคือ 2
                sheetIndex = 2;
            }

            // ------------------------
            // ใบต่อ (หน้า 2..n-1) – ใช้คอลัมน์เดียวกับหน้าแรก
            // ------------------------
            if (indexPage > 0) {
                data_last_page = model?.detail[key];

                addPage();
                Shot("i", form_background2, 0, 0, 210, 297);

                state.y = 10.5;
                // ใช้เลขแผ่นใหม่ รวม invoice + ใบปิด
                Shot("t", sheetIndex + ' / ' + totalSheets, CutPadding('w', 96.5), state.y, 'l'); //แผ่นที่ n/n
                state.y = 36.5;

                for (let item of (model?.detail[key]?.items || [])) {
                    // แถวบนสุด
                    Shot("t", item?.item_number || '', 18.5, state.y, 'c'); //รายการที่ (ใช้ X เดียวกับหน้าแรก)
                    Shot("t", item?.tariffifo?.code || '', CutPadding('w', 1.2), state.y - 12, 'l'); //ประเภทพิกัด
                    Shot("t", item?.cif_Value?.foreign || item?.price_foreign?.price_foreign || '', CutPadding('w', 30), state.y - 12, 'r'); //ราคาของ (เงินต่างประเทศ)
                    Shot("t", item?.duty?.duty_rate || '', CutPadding('w', 45), state.y - 12, 'r'); //อัตราอากรขาเข้า
                    Shot("t", item?.duty?.duty_fee || '', CutPadding('w', 59), state.y - 12, 'r'); // ค่าธรรมเนียม
                    Shot("t", item?.duty?.excise_code || '', CutPadding('w', 68.5), state.y - 12, 'c'); // รหัสสินค้าสรรพสามิต
                    Shot("t", item?.duty?.duty_excise || '', CutPadding('w', 86.5), state.y - 12, 'r'); // ภาษีสรรพสามิต
                    Shot("t", item?.duty?.duty_base_vat || '', CutPadding('w', 101.5), state.y - 12, 'r'); // ฐานภาษีมูลค่าเพิ่ม

                    // แถวกลาง
                    Shot("t", (item?.tariffifo?.sequence || item?.tariffifo?.statitic || '') + ' / ' + (item?.tariffifo?.sequence_unit || item?.tariffifo?.statitic_unit || ''), CutPadding('w', 3.9), state.y - 3.5, 'c'); //รหัสสถิติ หน่วย
                    Shot("t", (item?.cif_Value?.baht || item?.price_baht || ''), CutPadding('w', 30), state.y - 3.5, 'r'); //ราคาของ (บาท)
                    Shot("t", (item?.duty?.duty_amount || ''), CutPadding('w', 45.2), state.y - 3.5, 'r'); //อากรขาเข้าที่ชำระ
                    Shot("t", (item?.duty?.duty_other || ''), CutPadding('w', 59), state.y - 3.5, 'r'); //ภาษีอื่นๆ
                    Shot("t", (item?.duty?.excise_rate || ''), CutPadding('w', 72), state.y - 3.5, 'r'); //อัตราภาษีสรรพสามิต
                    Shot("t", (item?.duty?.duty_interior || ''), CutPadding('w', 87.5), state.y - 3.5, 'r'); //ภาษีเพื่อมหาดไทย
                    Shot("t", (item?.duty?.duty_vat || ''), CutPadding('w', 101.5), state.y - 3.5, 'r'); //ภาษีมูลค่าเพิ่ม

                    // แถวล่าง
                    Shot("t", (item?.privilege_code || ''), CutPadding('w', 2.3), state.y + 6, 'c'); //รหัสสิทธิพิเศษ
                    Shot("t", (item?.productifo?.net_weight?.weight || '') + ' ' + (item?.productifo?.net_weight?.unit_code || ''), CutPadding('w', 30), state.y + 6, 'r'); //น้ำหนักสุทธิ
                    Shot("t", (item?.quantity?.quantity || '') + ' ' + (item?.quantity?.UnitCode || ''), CutPadding('w', 45.1), state.y + 6, 'r'); //ปริมาณ

                    //ชนิดของ + คำอธิบาย (ใช้ X ตรงกับหน้าแรกอยู่แล้ว)
                    let reSizetxt_brand2 = resizeText(
                        (item?.productifo?.brandname || item?.productifo?.brand?.brand_name || ''),
                        50,
                        9
                    );
                    Shot("fs", reSizetxt_brand2);
                    Shot(
                        "t",
                        (item?.productifo?.brandname || item?.productifo?.brand?.brand_name || ''),
                        CutPadding('w', 51.5),
                        state.y + 2.5,
                        'l'
                    );

                    // รายละเอียดสินค้า: อังกฤษ / ไทย / ลักษณะที่ 1 / ลักษณะที่ 2 (จัดให้เหมือนหน้าแรก)
                    const engName2 = item?.goods_description?.english || item?.productifo?.english || '';
                    const thaiName2 = item?.goods_description?.thai || item?.productifo?.thai || '';
                    const attr1_2 = item?.productifo?.product_attribute1 || '';
                    const attr2_2 = item?.productifo?.product_attribute2 || '';

                    const descX2 = CutPadding('w', 46);
                    const baseY2 = state.y + 4.5;   // ใช้ offset เดียวกับหน้าแรก

                    // ชื่ออังกฤษ
                    const fsEng2 = resizeText(engName2, 92, 8);
                    Shot("fs", fsEng2);
                    Shot("t", engName2, descX2, baseY2, 'l');

                    // ชื่อไทย
                    const fsThai2 = resizeText(thaiName2, 120, 8);
                    Shot("fs", fsThai2);
                    Shot("t", thaiName2, descX2, baseY2 + 2, 'l');

                    // ลักษณะที่ 1
                    const fsAttr1_2 = resizeText(attr1_2, 90, 8);
                    Shot("fs", fsAttr1_2);
                    Shot("t", attr1_2, descX2, baseY2 + 4, 'l');

                    // ลักษณะที่ 2
                    const fsAttr2_2 = resizeText(attr2_2, 90, 8);
                    Shot("fs", fsAttr2_2);
                    Shot("t", attr2_2, descX2, baseY2 + 6, 'l');

                    // คืนฟอนต์กลับสำหรับคอลัมน์ด้านขวา (ประเทศ/สิทธิ์ภาษี ฯลฯ)
                    Shot("fs", 9);

                    Shot("t", (item?.origin_country_code || item?.productifo?.origin_country_code || ''), CutPadding('w', 102), state.y + 2.5, 'r');
                    Shot("t", (item?.import_tax_incentives_name || item?.productifo?.import_tax_incentives_name || ''), CutPadding('w', 102), state.y + 6.5, 'r');
                    Shot("t", (item?.import_tax_incentives_id || item?.productifo?.import_tax_incentives_id || ''), CutPadding('w', 102), state.y + 10.5, 'r');

                    //หมายเหตุ
                    Shot("t", (item?.remark || item?.productifo?.remark || ''), CutPadding('w', 51.5), state.y + 15.5, 'l');

                    //ใบอนุญาตินำเข้าหรือหนังสือรับรอง
                    Shot("t", (item?.license_number || (item?.permit && item?.permit[0]?.permit_no) || ''), CutPadding('w', 2), state.y + 14.5, 'l');

                    let index_various = 0;
                    for (let item2 of (item?.various_values || [])) {
                        Shot("t", (item2 || ''), CutPadding('w', 44.5), state.y + 11 + (2.5 * index_various), 'r');
                        index_various++;
                    }

                    state.y += 36; //อิงจาก เลขในรายการที่
                }

                // summary ใบต่อ (ใช้ X เหมือนหน้าแรก)
                state.y = 237.5;
                const pf2 = model?.detail[key]?.summary?.price_foreign;
                const pfText2 = pf2 ? `${pf2.price_foreign_amount || ''} ${pf2.price_foreign_unit || ''}`.trim() : '';
                Shot("t", pfText2 || '', CutPadding('w', 30), state.y, 'r'); //รวม ราคาของ (เงินต่างประเทศ)

                state.y = 248.5;
                Shot("t", (model?.detail[key]?.summary?.price_amount_baht || ''), CutPadding('w', 30), state.y, 'r'); //รวม ราคาของ (บาท)
                state.y = 242;
                Shot("fs", 15);
                Shot("t", (model?.detail[key]?.summary?.duty_amount || ''), CutPadding('w', 44.5), state.y, 'r'); // รวม ค่าอากรขาเข้า
                state.y = 245;
                Shot("fs", 10);
                const qty2 = model?.detail?.[key]?.summary?.total_quality;
                const qtyText2 = qty2 ? `${qty2.quantity_value || ''} ${qty2.quality_unit || ''}`.trim() : '';
                Shot("t", qtyText2 || '', CutPadding('w', 44.5), state.y, 'r'); // รวม ปริมาณ
                state.y = 241;
                Shot("fs", 14);
                Shot("t", (model?.detail[key]?.summary?.duty_fee || ''), CutPadding('w', 59), state.y, 'r'); // รวม ค่าธรรมเนียม
                state.y = 252.5;
                Shot("t", (model?.detail[key]?.summary?.other_duty || ''), CutPadding('w', 59), state.y, 'r'); // รวม ภาษีอื่นๆ

                state.y = 241;
                Shot("t", (model?.detail[key]?.summary?.excise_duty || ''), CutPadding('w', 87), state.y, 'r'); // รวม สรรพสามิต
                state.y = 252.5;
                Shot("t", (model?.detail[key]?.summary?.interior_duty || ''), CutPadding('w', 87), state.y, 'r'); // รวม ภาษีมหาดไทย
                state.y = 245.5;
                Shot("t", (model?.detail[key]?.summary?.vat_duty || ''), CutPadding('w', 102), state.y, 'r'); // รวม ภาษีมูลค่าเพิ่ม

                state.y = 258.5;
                Shot("t", (model?.detail[key]?.summary?.total_duty_amount || ''), CutPadding('w', 102), state.y, 'r'); // รวมค่าภาษีอากรทั้งสิ้น

                //footer ใบต่อ
                Shot("fs", 10);
                state.y = 259;
                Shot("t", (model?.detail[key]?.footer?.net_weight || model?.detail[key]?.footer?.total_net_weight?.net_weight_value || ''), CutPadding('w', 24.5), state.y, 'r'); // น้ำหนักสุทธิ
                state.y = 263;
                Shot("t", ('G.W. :' || ''), CutPadding('w', 3), state.y, 'l'); // G.W.
                Shot("t", (model?.detail[key]?.footer?.gross_weight || model?.detail[key]?.footer?.total_gross_weight?.gross_weight_value || ''), CutPadding('w', 22.5), state.y, 'r'); // G.W.

                state.y = 257;
                for (let item of (model?.detail[key]?.footer?.total_various_values || [])) {
                    Shot("t", (item || ''), CutPadding('w', 26), state.y, 'l');
                    state.y += 3;
                }

                state.y = 282;
                Shot("t", (model?.detail[key]?.footer?.status || ''), CutPadding('w', -5), state.y, 'l'); // สถานะล่าสุด
                state.y = 286;
                Shot("t", (model?.detail[key]?.footer?.status_date || ''), CutPadding('w', -5), state.y, 'l'); // วันที่สถานะล่าสุด

                state.y = 279;
                Shot("t", (model?.detail[key]?.footer?.signature_name || model?.detail[key]?.footer?.signature?.signature_name || ''), CutPadding('w', 64), state.y, 'l'); // ลายมือชื่ออิเล็กทรอนิกส์
                state.y = 284;
                Shot("t", (model?.detail[key]?.footer?.signature_date || ''), CutPadding('w', 64), state.y, 'l'); // วันที่ยื่น

                // จบใบต่อหนึ่งหน้า -> แผ่นถัดไป
                sheetIndex++;
            }

            indexPage++;
        }

        // =========================
        // หน้ารวม Invoice (page_invoice)
        // =========================
        if (model.detail?.page_invoice?.invoice && model.detail.page_invoice.invoice.length > 0) {
            addPage();
            Shot("i", form_background2, 0, 0, 210, 297);
            state.y = 10.5;
            // ใช้ sheetIndex / totalSheets
            Shot("t", sheetIndex + ' / ' + totalSheets, CutPadding('w', 96.5), state.y, 'l');
            let txt_invoice = '';
            let index_invoice = 0;

            for (let item of (model.detail?.page_invoice?.invoice || [])) {
                if (index_invoice == model.detail?.page_invoice?.invoice.length - 1) {
                    txt_invoice += item;
                } else {
                    txt_invoice += item + ', ';
                }
                index_invoice++;
            }
            let cutTextInvoice = splitText(txt_invoice, CutPadding('w', 100));
            state.y = 56.5;

            Shot("t", 'Invoice No. :', CutPadding('w', 2), state.y, 'l');
            for (let item2 of cutTextInvoice) {
                Shot("t", item2, CutPadding('w', 9.5), state.y, 'l');
                state.y += 3;
            }

            //summary ใช้ data_last_page
            state.y = 237.5;
            Shot("t", (data_last_page?.summary?.price_amount || data_last_page?.summary?.price_foreign?.price_foreign_amount || ''), CutPadding('w', 30), state.y, 'r'); //รวม ราคาของ (เงินต่างประเทศ)
            state.y = 248.5;
            Shot("t", (data_last_page?.summary?.price_amount_baht || ''), CutPadding('w', 30), state.y, 'r'); //รวม ราคาของ (บาท)
            state.y = 242;
            Shot("fs", 15);
            Shot("t", (data_last_page?.summary?.duty_amount || ''), CutPadding('w', 44.5), state.y, 'r'); // รวม ค่าอากรขาเข้า
            state.y = 245;
            Shot("fs", 10);
            Shot("t", (data_last_page?.summary?.quantity_amount || data_last_page?.summary?.total_quality?.quantity_value || ''), CutPadding('w', 44.5), state.y, 'r'); // รวม ปริมาณ
            state.y = 241;
            Shot("fs", 14);
            Shot("t", (data_last_page?.summary?.duty_fee || ''), CutPadding('w', 59), state.y, 'r'); // รวม ค่าธรรมเนียม
            state.y = 252.5;
            Shot("t", (data_last_page?.summary?.other_duty || ''), CutPadding('w', 59), state.y, 'r'); // รวม ภาษีอื่นๆ

            state.y = 241;
            Shot("t", (data_last_page?.summary?.excise_duty || ''), CutPadding('w', 87), state.y, 'r'); // รวม สรรพสามิต
            state.y = 252.5;
            Shot("t", (data_last_page?.summary?.interior_duty || ''), CutPadding('w', 87), state.y, 'r'); // รวม ภาษีมหาดไทย
            state.y = 245.5;
            Shot("t", (data_last_page?.summary?.vat_duty || ''), CutPadding('w', 102), state.y, 'r'); // รวม ภาษีมูลค่าเพิ่ม

            state.y = 258.5;
            Shot("t", (data_last_page?.summary?.total_duty_amount || ''), CutPadding('w', 102), state.y, 'r'); // รวมค่าภาษีอากรทั้งสิ้น

            //footer
            Shot("fs", 10);
            state.y = 259;
            Shot("t", (data_last_page?.footer?.net_weight || data_last_page?.footer?.total_net_weight?.net_weight_value || ''), CutPadding('w', 24.5), state.y, 'r'); // น้ำหนักสุทธิ
            state.y = 263;
            Shot("t", ('G.W. :' || ''), CutPadding('w', 3), state.y, 'l'); // G.W.
            Shot("t", (data_last_page?.footer?.gross_weight || data_last_page?.footer?.total_gross_weight?.gross_weight_value || ''), CutPadding('w', 22.5), state.y, 'r'); // G.W.

            state.y = 257;
            for (let item of (data_last_page?.footer?.total_various_values || [])) {
                Shot("t", (item || ''), CutPadding('w', 26), state.y, 'l');
                state.y += 3;
            }

            state.y = 282;
            Shot("t", (data_last_page?.footer?.status || ''), CutPadding('w', -5), state.y, 'l'); // สถานะล่าสุด
            state.y = 286;
            Shot("t", (data_last_page?.footer?.status_date || ''), CutPadding('w', -5), state.y, 'l'); // วันที่สถานะล่าสุด

            state.y = 279;
            Shot("t", (data_last_page?.footer?.signature_name || data_last_page?.footer?.signature?.signature_name || ''), CutPadding('w', 64), state.y, 'l'); // ลายมือชื่ออิเล็กทรอนิกส์
            state.y = 284;
            Shot("t", (data_last_page?.footer?.signature_date || ''), CutPadding('w', 64), state.y, 'l'); // วันที่ยื่น

            // จบหน้า invoice -> แผ่นถัดไป
            sheetIndex++;
        }

        // =========================
        // ใบปิด (หน้า final)
        // =========================
        addPage();
        Shot("i", form_background3, 0, 0, 210, 297);
        state.y = 15.5;
        // ใช้ sheetIndex เป็นแผ่นสุดท้าย
        Shot("t", sheetIndex + ' / ' + totalSheets, CutPadding('w', 96.7), state.y, 'l'); //แผ่นที่ n/n
        state.y = 16.5;
        Shot("fs", 16);
        Shot("t", 'เลขที่ใบขน ' + (model?.document_control?.reference_number || ''), CutPadding('w', 54), state.y, 'l'); //เลขที่ใบขน

        // สำหรับผู้นำของเข้า
        state.y = 30;
        Shot("fs", 12);
        if (model?.endorse?.inspection_request) {
            Shot("t", 'Inspection Request', CutPadding('w', 3), state.y, 'l'); //Inspection Request
            Shot("t", model?.endorse?.inspection_request || '', CutPadding('w', 23), state.y, 'l'); //Inspection Request
        }

        if (model?.endorse?.reassessment_request) {
            state.y += 6;
            Shot("t", 'Reassessment Request', CutPadding('w', 3), state.y, 'l'); //Reassessment Request
            Shot("t", model?.endorse?.reassessment_request || '', CutPadding('w', 23), state.y, 'l'); //Reassessment Request
        }
        if (model?.endorse?.cargo_packing_type) {
            state.y += 6;
            Shot("t", 'Cargo Packing Type', CutPadding('w', 3), state.y, 'l'); //Cargo Packing Type
            Shot("t", model?.endorse?.cargo_packing_type || '', CutPadding('w', 23), state.y, 'l'); //Cargo Packing Type
        }

        // บันทึกการปล่อย
        if (model?.endorse?.seal_date) {
            state.y = 236;
            Shot("t", 'Seal Date', CutPadding('w', 3), state.y, 'l'); //Seal Date

            state.y += 6;
            Shot("t", 'Release Date' || '', CutPadding('w', 3), state.y, 'l'); //Release Date
            Shot("t", model?.endorse?.seal_date?.release_date || '', CutPadding('w', 16), state.y, 'l'); //Release Date
            state.y += 6;
            Shot("t", 'Delivery Date' || '', CutPadding('w', 3), state.y, 'l'); //Delivery Date
            Shot("t", model?.endorse?.seal_date?.delivery_date || '', CutPadding('w', 16), state.y, 'l'); //Delivery Date
        }

        // ==== วาด send_acc แนวตั้งทุกหน้า ยกเว้นหน้าสุดท้าย ==== 
        const sendAcc = model?.send_acc || model?.detail?.send_acc || '';

        if (sendAcc) {
            const totalPages = doc.internal.getNumberOfPages();

            // ============ OFFSET ของแต่ละหน้า =============
            // ถ้ายังไม่รู้ ให้ลองปรินต์ตัวเลขแล้วผมจะปรับให้
            const pageOffset = {
                1: { x: 15.5, y: 266 },
                2: { x: 14.5, y: 255 },
                3: { x: 14.5, y: 255 },

            };

            const step = 1.1;   // ระยะห่างระหว่างตัวอักษร
            const angle = 90;   // หมุน 90 องศา

            for (let page = 1; page <= totalPages - 1; page++) {
                doc.setPage(page);
                Shot('sf', 'CordiaNew', 'normal', 9);

                const px = pageOffset[page]?.x ?? pageOffset[1].x;
                let py = pageOffset[page]?.y ?? pageOffset[1].y;

                for (let i = 0; i < sendAcc.length; i++) {
                    const ch = sendAcc[i];
                    Shot("rt", ch, px, py, 'l', angle);
                    py -= step;
                }
            }
        }
        // ==== เพิ่มเลขหน้าให้หน้าแรก (เช่น 1/4) ====
        {
            const totalPages = doc.internal.getNumberOfPages();  // ตอนนี้สร้างครบทุกหน้าแล้ว
            doc.setPage(1);                                     // กลับไปหน้า 1
            Shot('sf', 'CordiaNew', 'normal', 10);              // ตั้งฟอนต์ให้เหมือนหน้าอื่น
            Shot('fs', 10);

            const pageText = `1 / ${totalPages}`;               // เช่น 1 / 4
            const pageY = 16.5;                                 // ใช้ Y เดียวกับหน้าอื่น
            Shot('t', 'แผ่นที่ ' + pageText, CutPadding('w', 86.5), pageY, 'l');
        }

        const filename = req?.query?.file_name || "ใบขนสินค้าขาเข้า"; // ชื่อไฟล์เป็นภาษาไทย
        const buffer = await Shot("save", filename);

        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.setHeader("Content-Type", "application/pdf");
        res.send(buffer);

    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});





router.post('/delivery-transfer-notes', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;




        // ใช้งาน:
        const form_background = await getBase64Image('./app/form_report/delivery-transfer-notes_f.jpg');
        const form_background2 = await getBase64Image('./app/form_report/delivery-transfer-notes_b.jpg');
        function textseting(keytext) {
            return keytext ? (keytext).toString() : ''
        }

        // console.log('base64',base64);
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },


            ]
        });
        const { Shot, CutPadding, state, addPage, splitText } = pdf;
        form_background2

        let list = Object.keys(model.detail)
            .filter(key => /^page_\d+$/.test(key))
            .sort((a, b) => Number(a.split("_")[1]) - Number(b.split("_")[1]))
            .map(key => model.detail[key]);
        console.log(list);



        for (let p = 0; p < list.length; p++) {
            const page = list[p];
            Shot('sf', 'CordiaNew', 'bold', 12);
            if (p > 0) {
                addPage();
            }
            Shot("i", p > 0 ? form_background2 : form_background, 0, 0, 210, 297);

            Shot("i", model.barcode_base64, 8, 7, 50, 10);
            state.y = 38
            Shot("tb", textseting(p + 1), CutPadding('w', 100, 1), state.y - 1, 'l', 2);
            Shot("tb", textseting(model.reference_number), CutPadding('w', 10, 1), state.y, 'l', 2);
            Shot("tb", textseting(model.goods_transition_number), CutPadding('w', 65, 2), state.y, 'l', 2);
            state.y += 5.5
            Shot('sf', 'CordiaNew', 'normal', 12);

            Shot("tb", textseting(model.company_info.thai_name), CutPadding('w', 1, 2), state.y, 'l', 2);
            Shot("tb", textseting(model.company_info.tax_id), CutPadding('w', 88, 2), state.y, 'l', 2);

            state.y += 5.5
            Shot("tb", textseting(model.company_info.origin_address), CutPadding('w', 0, 0), state.y, 'l', 2);
            Shot('sf', 'CordiaNew', 'bold', 12);
            state.y += 10
            Shot("tb", textseting(model.transport_information.indicator_container == 'Y' ? 'x' : ''), CutPadding('w', 0, 1.2), state.y + 1, 'l', 2);
            Shot("tb", textseting(model.transport_information.container_type), CutPadding('w', 29,), state.y, 'l', 2);
            Shot("tb", textseting(model.transport_information.container_no), CutPadding('w', 58,), state.y, 'l', 2)
            Shot("tb", textseting(model.transport_information.container_size_type_code), CutPadding('w', 95,), state.y, 'c', 2)
            state.y += 6

            Shot("tb", textseting(model.transport_information.indicator_bulk == 'Y' ? 'x' : ''), CutPadding('w', 0, 1.5), state.y + 1, 'l', 2);
            Shot('sf', 'CordiaNew', 'normal', 12);
            Shot("tb", textseting(model.transport_information.car_type), CutPadding('w', 27,), state.y, 'l', 2);
            Shot("tb", textseting(model.transport_information.car_license), CutPadding('w', 60,), state.y, 'l', 2)
            Shot("tb", textseting(model.transport_information.car_province_code), CutPadding('w', 95,), state.y, 'c', 2)
            state.y += 5.2

            Shot("tb", textseting(model.transport_information.trailer_license), CutPadding('w', 60,), state.y, 'l', 2)
            Shot("tb", textseting(model.transport_information.trailer_province_code), CutPadding('w', 95,), state.y, 'c', 2)
            state.y += 5

            Shot("tb", textseting(model.transport_information.driver_name), CutPadding('w', 3,), state.y, 'l', 2)
            Shot("tb", textseting(model.transport_information.contact_number), CutPadding('w', 85,), state.y, 'l', 2)
            state.y += 5.7

            Shot("tb", textseting(model.transport_information.reefer_temperature), CutPadding('w', 12,), state.y, 'l', 2)
            Shot("tb", textseting(model.transport_information.ventilation), CutPadding('w', 47,), state.y, 'c', 2)
            state.y += 5.7
            Shot('sf', 'CordiaNew', 'bold', 12);
            Shot("tb", `${textseting(model.transport_information.release_port.release_port_code)} ${textseting(model.transport_information.release_port.release_port_name)}`, CutPadding('w', 26,), state.y, 'l', 2)

            state.y += 5.7
            Shot("tb", `${textseting(model.transport_information.discharge_port.discharge_port_code)} ${textseting(model.transport_information.discharge_port.discharge_port_name)}`, CutPadding('w', 9,), state.y, 'l', 2)

            state.y += 5.7
            Shot("tb", `${textseting(model.transport_information.packing_port.packing_port_code)} ${textseting(model.transport_information.packing_port.packing_port_name)}`, CutPadding('w', 9,), state.y, 'l', 2)

            state.y += 5.7
            Shot("tb", `${textseting(model.transport_information.vessel_name)}`, CutPadding('w', 9,), state.y, 'l', 2)
            Shot('sf', 'CordiaNew', 'normal', 12);
            Shot("tb", `${textseting(model.transport_information.voyage_number)}`, CutPadding('w', 94,), state.y, 'l', 2)
            Shot('sf', 'CordiaNew', 'bold', 12);
            state.y += 5.7
            Shot("tb", `${textseting(model.transport_information.loading_port.loading_port_code)} ${textseting(model.transport_information.loading_port.loading_port_name)}`, CutPadding('w', 9,), state.y, 'l', 2)

            Shot("tb", `${textseting(model.transport_information.departure_date)}`, CutPadding('w', 92,), state.y, 'l', 2)

            state.y += 5.7
            /*  console.log(CutPadding('w', 83)); */
            Shot('sf', 'CordiaNew', 'normal', 12);
            let cuting = splitText(`${model.transport_information.remark}`, CutPadding('w', 83));
            /*  console.log(cuting, `${model.transport_information.remark}`); */

            for (let item2 of cuting) {
                Shot("tb", `${textseting(item2)}`, CutPadding('w', 9,), state.y, 'l', 2)
                state.y += 5.7
            }

            state.y = 134
            let len = 7

            for (let index = 0; index < page.item.length; index++) {
                const i = page.item[index];
                console.log(i);

                Shot("tb", `${textseting(i.detail_number)}`, CutPadding('w', 0, -9), state.y + 1, 'l', 2)
                Shot('sf', 'CordiaNew', 'bold', 12);
                Shot("tb", `${textseting(i.document_declaration.declaration_number)}`, CutPadding('w', 0), state.y, 'l', 3)
                Shot('sf', 'CordiaNew', 'normal', 12);
                Shot("tb", `M: ${textseting(i.document_declaration.master_bill_Lading)}`, CutPadding('w', 0), state.y + 3.6, 'l', 2)
                Shot("tb", `H: ${textseting(i.document_declaration.house_bil_lading)}`, CutPadding('w', 0), state.y + 7.2, 'l', 2)

                if (i.document_declaration?.UNNO) {
                    Shot("tb", `UN NO : ${textseting(i.document_declaration.UNNO)}`, CutPadding('w', 0), state.y + 14.4, 'l', 2)
                }
                if (i.document_declaration?.IMOCLASS) {
                    Shot("tb", `IMO CLASS : ${textseting(i.document_declaration.IMOCLASS)}`, CutPadding('w', 0), state.y + 18, 'l', 2)
                }
                let textall = `${textseting(i.exporter.thai_name)}\n${textseting(i.exporter.english_name)}\n${textseting(i.exporter.address)} ${textseting(i.exporter.contact_number)} (${textseting(i.exporter.tax_id)}) `
                let cuting = splitText(`${textall}`, CutPadding('w', 30));
                let font_size = 12;
                let max_line = 7;
                let new_font = max_line / cuting.length * font_size;
                let line_height = new_font / 3.6;

                if (new_font > 12) {
                    new_font = 12;
                }
                if (line_height > 3.6) {
                    line_height = 3.6;
                }

                Shot("fs", new_font);

                cuting = splitText(`${textall}`, CutPadding('w', 30));

                //  console.log(cutText); 

                let sumpos = 0
                for (let item2 of cuting) {
                    Shot("t", item2, CutPadding('w', 24, 1), state.y + sumpos, 'l');
                    sumpos += line_height
                }

                /* let cutText = splitText(`${textseting(i.exporter.thai_name)}\n${textseting(i.exporter.english_name)}\n${textseting(i.exporter.address)} ${textseting(i.exporter.contact_number)} (${textseting(i.exporter.tax_id)}) `, CutPadding('w', 32));
                let sumpos = 0
                for (let c = 0; c < cutText.length; c++) {
                    const ci = cutText[c];
                    Shot("t", `${ci}`, CutPadding('w', 24, 1), state.y + sumpos, 'l')
                    sumpos += 3.6

                } */
                Shot('sf', 'CordiaNew', 'bold', 12);
                Shot("tb", `${textseting(i.total_package.package_amount)} ${textseting(i.total_package.package_unit)}`, CutPadding('w', 85), state.y, 'r', 2)
                Shot("tb", `${textseting(i.total_gross_weight.gross_weight)} ${textseting(i.total_gross_weight.gross_weight_unit)}`, CutPadding('w', 100, 12), state.y, 'r', 2)
                Shot('sf', 'CordiaNew', 'normal', 12);
                if (i.reference_number) {
                    Shot("tb", `REF :${textseting(i.reference_number)}`, CutPadding('w', 0), state.y + 22.5, 'l', 2)
                }
                if (i.job_code) {
                    Shot("tb", `job :${textseting(i.job_code)}`, CutPadding('w', 0), state.y + 25.5, 'l', 2)
                }
                if (i.invoice_number) {
                    Shot("tb", `INV :${textseting(i.invoice_number)}`, CutPadding('w', 24, 1), state.y + 22.5, 'l', 2)
                }



                state.y += 31
            }
            let ispage = p > 0 ? 125 : 94
            state.y = 134 + ispage
            if (page.summary) {
                Shot('sf', 'CordiaNew', 'bold', 12);
                if (page.summary.package_amount) {
                    Shot("tb", `${textseting(page.summary.package_amount.package_amount_value)} ${textseting(page.summary.package_amount.package_amount_unit)}`, CutPadding('w', 85), state.y, 'r', 2)

                }
                if (page.summary.gross_weight) {
                    Shot("tb", `${textseting(page.summary.gross_weight.gross_weight_value)} ${textseting(page.summary.gross_weight.gross_weight_unit)}`, CutPadding('w', 100, 12), state.y, 'r', 2)


                }
            }
            state.y = 244
            if (page.footer) {

                Shot('sf', 'CordiaNew', 'normal', 16);
                Shot("tb", `SEND TIME : `, CutPadding('w', 0, -14.5), state.y, 'l', 2)
                Shot('sf', 'CordiaNew', 'normal', 12);
                Shot("tb", `${page.footer.send_time}`, CutPadding('w', 14, -14.5), state.y, 'l', 2)

                Shot('sf', 'CordiaNew', 'normal', 16);
                Shot("tb", `,RECEIVE TIME :`, CutPadding('w', 20,), state.y, 'l', 2)
                Shot('sf', 'CordiaNew', 'normal', 12);
                Shot("tb", `${page.footer.receive_time}`, CutPadding('w', 38), state.y, 'l', 2)

                Shot('sf', 'CordiaNew', 'bold', 12);
                Shot("tb", `${page.footer.booking_number}`, CutPadding('w', 11), state.y + 6, 'l', 2)
                Shot("tb", `${page.footer.vgm_weight.vgm_weight_value} ${page.footer.vgm_weight.vgm_weight_unit}`, CutPadding('w', 7), state.y + 13, 'l', 2)
                Shot("tb", `${page.footer.vgm_authorized_name}`, CutPadding('w', 19), state.y + 21, 'l', 2)
                Shot("tb", `${page.footer.mode_movment}`, CutPadding('w', 13), state.y + 28, 'l', 2)
                Shot('sf', 'CordiaNew', 'normal', 12);
                Shot("tb", `${page.footer.signature_name}`, CutPadding('w', 90), state.y + 11, 'c', 2)
                Shot("tb", `${page.footer.signature_date}`, CutPadding('w', 90), state.y + 15, 'c', 2)
                Shot('sf', 'CordiaNew', 'bold', 12);
                let cut = splitText(`${textseting(page.footer.container_operator.container_operator_name)}\n(${textseting(page.footer.container_operator.container_operator_tax_id)})`, CutPadding('w', 40));
                let pos = 28
                for (let c = 0; c < cut.length; c++) {
                    const ci = cut[c];
                    Shot("tb", `${ci}`, CutPadding('w', 50, 1), state.y + pos, 'l', 2)
                    pos += 3.6

                }
            }



        }

        const filename = req?.query?.file_name || "แบบฟอร์มชำระค่าภาษี"; // ชื่อไฟล์เป็นภาษาไทย
        console.log('filename');
        const buffer = await Shot("save", filename);
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});


router.post('/formFee50', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;




        // ใช้งาน:
        const form_background = await getBase64Image('./app/form_report/formVat10_50.jpg');


        // console.log('base64',base64);
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },


            ]
        });
        const { Shot, CutPadding, state, addPage, resizeText } = pdf;
        let tostart = 0
        for (let index = 0; index < model.length; index++) {

            let i = model[index]

            Shot("i", form_background, 0, 0 + tostart, 210, 297);


            Shot("i", i.qrcode_base64, 18, 28 + tostart, 30, 30);//qr_code
            Shot("i", i.barcode_base64, 55, 28 + tostart, 130, 18);//bar_code
            state.y = 14 + tostart;
            //Shot("fs", 16);
            Shot('sf', 'CordiaNew', 'bold', 16);
            Shot("tb", i.declaration_number || '', CutPadding('w', 88), state.y, 'l', 2);
            Shot('sf', 'CordiaNew', 'bold', 16);
            state.y = 74 + tostart;
            let reSizetxt = resizeText(i.name || '', 175, 16);
            Shot("fs", reSizetxt);
            Shot("tb", i.name || '', CutPadding('w', 6), state.y, 'l', 1);
            Shot('sf', 'CordiaNew', 'bold', 16);
            state.y += 6.5
            Shot("tb", i.ref_1 || '', CutPadding('w', 33), state.y, 'l', 2);
            state.y += 6
            Shot("tb", i.ref_2 || '', CutPadding('w', 33), state.y, 'l', 2);
            state.y += 6
            Shot("tb", i.price || '', CutPadding('w', 20), state.y, 'l', 2);


            /*   Shot("t", model.declaration_number || '', CutPadding('w', 50), state.y, 'c');
      
      
      
              state.y = 87.5;
      
              Shot("t", model.name || '', CutPadding('w', 6), state.y, 'l');
      
              state.y += 10
      
              Shot("t", model.ref_1 || '', CutPadding('w', 33), state.y, 'l');
      
      
              state.y += 10
      
              Shot("t", model.ref_2 || '', CutPadding('w', 33), state.y, 'l');
      
      
              state.y += 10
      
              Shot("t", model.price || '', CutPadding('w', 20), state.y, 'l'); */

            if (tostart == 0) {
                tostart = 149
            } else {
                if ((model.length - 1) > index) {
                    addPage();
                    tostart = 0
                }

            }
        }




        const filename = req?.query?.file_name || "แบบฟอร์มชำระค่าภาษี"; // ชื่อไฟล์เป็นภาษาไทย
        console.log('filename');
        const buffer = await Shot("save", filename);
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});


router.post('/airwayReport', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {


        let model = req.body;
        console.log(model);

        const form_background = await getBase64Image('./app/form_report/airwayReport.jpg');
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                /*   { name: 'THSarabunNew', vfsName: 'THSarabunNew.ttf', data: Get_font_pdf_th2(), style: 'bold' }, */
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },

            ]
        });
        const { Shot, CutPadding, state, splitText, resizeText } = pdf;
        Shot("i", form_background, 0, 0, 210, 297);
        Shot('sf', 'CordiaNew', 'normal', 10);

        Shot("t", model.reference_number, 37, 28.5, 'l')
        Shot("i", model.airway_bill_number, 122.5, 21, 53, 14);

        Shot('sf', 'CordiaNew', 'normal', 10);
        let reSizethainame = resizeText(model.applicant_information.thai_name, 88, 10);
        Shot('sf', 'CordiaNew', 'normal', reSizethainame);
        Shot("t", model.applicant_information.thai_name, 35, 38.5, 'l')


        Shot('sf', 'CordiaNew', 'normal', 10);
        Shot("t", model.applicant_information.tax_id + ' สาขา ' + model.applicant_information.branch, 159, 38.5, 'l')
        let reSizeaddress = resizeText(model.applicant_information.address, 175, 10);
        Shot('sf', 'CordiaNew', 'normal', reSizeaddress);
        Shot("t", model.applicant_information.address, 23, 46.5, 'l')


        Shot('sf', 'CordiaNew', 'normal', 10);
        let reSizeairport = resizeText(model.transport_information.airport, 82, 10);
        Shot("t", model.transport_information.airport, 35, 55, 'l')
        Shot('sf', 'CordiaNew', 'normal', reSizeairport);

        Shot("sf", 'CordiaNew', 'normal', 10);
        Shot("t", model.transport_information.flight_number, 17, 79, 'l')
        Shot("t", model.transport_information.flight_date, 42, 79, 'l')
        Shot("i", model.transport_information.bill_lading.master_number, 76, 73.5, 32, 14.5);
        Shot("t", model.cargo_information.warehouse, 137, 79, 'l');
        Shot("t", model.cargo_information.package_amount, 158.5, 79, 'l');
        Shot("t", model.cargo_information.gross_weight, 179, 79, 'l');
        Shot("i", model.transport_information.bill_lading.houser_number, 76, 90.5, 35, 11.5);
        let resizedeclarant = resizeText(model.submission_information.declarant, 73, 10);
        Shot("t", model.submission_information.declarant, 42.5, 122, 'l');
        Shot('sf', 'CordiaNew', 'normal', resizedeclarant);

        Shot('sf', 'CordiaNew', 'normal', 10);
        Shot("t", model.submission_information.submission_date + '  ' + model.submission_information.submission_time, 44, 130, 'l');
        Shot("t", 'Invoice no. ' + model.submission_information.invoice_no + ' ' + model.submission_information.invoice_date, 17, 149, 'l');




        const buffer = await Shot("save", 'ลองเปลี่ยนดู');
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});


router.post('/nameChangeCertificate', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {


        let model = req.body;
        console.log(model);

        const form_background = await getBase64Image('./app/form_report/nameChangeCertificate.jpg');
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                /*   { name: 'THSarabunNew', vfsName: 'THSarabunNew.ttf', data: Get_font_pdf_th2(), style: 'bold' }, */
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },
                { name: 'tahoma', vfsName: 'tahoma Normal.ttf', data: stripDataUrl(tahomaNormal), style: 'normal' },
            ]
        });


        const { Shot, CutPadding, state, splitText, resizeText } = pdf;
        const checkImg = await getBase64Image('./app/form_report/check.png');
        const HEADER_SIZE = 12;
        const BODY_SIZE = 10;
        const ROW = 5;
        const T = v => String(v ?? '');
        const drawCheck = (flag, x, y) => {
            if (flag === 'Y') {
                // วาดเฉพาะตอนเป็น Y, ถ้า N ก็ปล่อยว่าง
                Shot("i", checkImg, x, y, 5, 5); // ปรับ size ให้พอดีกล่อง
            }
        };


        Shot("i", form_background, 0, 0, 210, 297);
        Shot('sf', 'tahoma', 'normal', HEADER_SIZE);

        Shot("t", String(model?.page_no ?? ''), 187, 22, 'l');
        Shot('sf', 'tahoma', 'normal', BODY_SIZE);

        /*let reSizeaddress = resizeText( model.applicant_information.address, 175, 10);
        Shot('sf', 'CordiaNew', 'normal', reSizeaddress);
        Shot("t", model.applicant_information.address, 23, 46.5, 'l') */

        let resizeinfothainame = resizeText(model.applicant_information.thai_name, 96, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeinfothainame);
        Shot("t", model.applicant_information.thai_name, 30.8, 35, 'l')

        let resizeinfotaxid = resizeText(model.applicant_information.tax_id, 36, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeinfotaxid);
        Shot("t", model.applicant_information.tax_id, 163, 35, 'l')
        /* Shot('sf', 'CordiaNew', 'normal', 15.5); */

        let resizeAddress = resizeText(model.applicant_information.address, 180, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeAddress);
        Shot("t", model.applicant_information.address, 20, 41.5, 'l')


        let resizeOldname = resizeText(model.transport_information.old_vessel_name, 103, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeOldname);
        Shot("t", model.transport_information.old_vessel_name, 40.5, 47.8, 'l')

        Shot("t", model.transport_information.old_voyage_number, 160, 47.8, 'l')

        drawCheck(model?.transport_information?.new_vessel_voyage?.indicator_vessel_voyage, 10, 56.5);

        drawCheck(model?.transport_information?.release_port?.indicator_release_port, 10, 63);

        let resizeLoadport = resizeText(model.transport_information.load_port.port_name + ' ' + model.transport_information.load_port.port_code, 160, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeLoadport);
        Shot("t", model.transport_information.load_port.port_name + ' ' + model.transport_information.load_port.port_code, 40.5, 54.5, 'l')

        let resizeNewvesselnname = resizeText(model.transport_information.new_vessel_voyage.vessel_name, 66, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeNewvesselnname);
        Shot("t", model.transport_information.new_vessel_voyage.vessel_name, 78.5, 60.5, 'l')

        Shot('sf', 'tahoma', 'normal', BODY_SIZE);
        Shot("t", model.transport_information.new_vessel_voyage.voyage_number, 159.5, 60.5, 'l')

        let resizePortcode = resizeText(model.transport_information.release_port.port_name + ' ' + model.transport_information.release_port.port_code, 130, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizePortcode);
        Shot("t", model.transport_information.release_port.port_name + ' ' + model.transport_information.release_port.port_code, 70.5, 66.8, 'l')


        Shot('sf', 'tahoma', 'normal', BODY_SIZE);

        let y = 85;
        for (const item of model.detail) {
            const d = item?.container_information ?? {};

            const col_detail = 16;   // เลขลำดับ
            const col_shipInv = 26;   // เลขใบกำกับการขนย้ายสินค้า
            const col_contNo = 83;   // หมายเลขตู้คอนเทนเนอร์
            const col_pkg = 160;  // จำนวนหีบห่อ/หน่วย
            const col_gw = 198;  // น้ำหนักรวมหีบห่อ/หน่วย

            // ลำดับ
            Shot('sf', 'tahoma', 'normal', BODY_SIZE);
            Shot("t", T(item?.detail_namber), col_detail, y, 'l');

            // Shipping Invoice (ปรับขนาดเฉพาะตัวนี้)
            const invSize = resizeText(T(d?.shipping_invoice_number), 50, BODY_SIZE);
            Shot('sf', 'tahoma', 'normal', invSize);
            Shot("t", T(d?.shipping_invoice_number), col_shipInv, y, 'l');

            // Container No (ปรับขนาด)
            const contSize = resizeText(T(d?.container_no), 45, BODY_SIZE);
            Shot('sf', 'tahoma', 'normal', contSize);
            Shot("t", T(d?.container_no), col_contNo, y, 'l');
            // จำนวน + หน่วย (ปรับขนาด และชิดขวาตามกรอบ)
            const pkgText = `${T(d?.package_amount)} ${T(d?.package_unit)}`.trim();
            const pkgSize = resizeText(pkgText, 30, BODY_SIZE);
            Shot('sf', 'tahoma', 'normal', pkgSize);
            Shot("t", pkgText, col_pkg, y, 'r');

            // Gross weight + unit (ปรับขนาด และชิดขวา)
            const gwText = `${T(d?.gross_weight)} ${T(d?.gross_weight_unit)}`.trim();
            const gwSize = resizeText(gwText, 32, BODY_SIZE);
            Shot('sf', 'tahoma', 'normal', gwSize);
            Shot("t", gwText, col_gw, y, 'r');

            // ขยับลงแถวถัดไป
            y += ROW;
        }


        /* Shot("i", model.submission_information.representative_signature, 157, 269, 25, 10, 'l'); */
        Shot("t", model.submission_information.submission_date, 160, 287, 'l');


        const buffer = await Shot("save", 'ลองเปลี่ยนดู');
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});

router.post('/tax-compensation-report', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {


        let model = req.body;
        console.log(model);

        const form_background = await getBase64Image('./app/form_report/tax_compensation.jpg');
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                /*   { name: 'THSarabunNew', vfsName: 'THSarabunNew.ttf', data: Get_font_pdf_th2(), style: 'bold' }, */
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },
                { name: 'tahoma', vfsName: 'tahoma Normal.ttf', data: stripDataUrl(tahomaNormal), style: 'normal' },
            ]
        });

        const { Shot, CutPadding, state, splitText, resizeText } = pdf;

        const HEADER_SIZE = 13;
        const BODY_SIZE = 10;




        Shot("i", form_background, 0, 0, 210, 297);


        Shot('sf', 'tahoma', 'normal', HEADER_SIZE);
        Shot('t', model.compensation_code, 112, 55.2, 'l');

        let resizeappNum = resizeText(model.application_number, 43, HEADER_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeappNum);
        Shot('t', model.application_number, 153, 15.5, 'l');

        let resizeComname = resizeText(model.applicant.company_name, 135, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeComname);
        Shot('t', model.applicant.company_name, 59.5, 79.5, 'l');

        let resizeAddress = resizeText(model.applicant.address, 158, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeAddress);
        Shot('t', model.applicant.address, 33.8, 87.5, 'l');

        let resizeTaxid = resizeText(model.applicant.tax_id, 84, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeTaxid);
        Shot('t', model.applicant.tax_id, 54, 95.5, 'l');

        let resizeTel = resizeText(model.applicant.telephone, 30, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeTel);
        Shot('t', model.applicant.telephone, 158, 95.5, 'l');

        let resizeRepname = resizeText(model.applicant.representative.name, 117, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeRepname);
        Shot('t', model.applicant.representative.name, 20, 103.5, 'l');

        let resizePosition = resizeText(model.applicant.representative.position, 36, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizePosition);
        Shot('t', model.applicant.representative.position, 156, 103.5, 'l');

        let resizeAge = resizeText(model.applicant.representative.age, 10, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeAge);
        Shot('t', String(model?.applicant?.representative?.age ?? ''), 20, 111.2, 'l');

        let resizeNationality = resizeText(model.applicant.representative.nationality, 30, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeNationality);
        Shot('t', model.applicant.representative.nationality, 100, 111.2, 'l');

        let resizeDeclarant = resizeText(model.export_information.export_declaration_no, 56, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeDeclarant);
        Shot('t', model.export_information.export_declaration_no, 137, 121.5, 'l');

        let resizeQuantity = resizeText(model.export_information.quantity_exported, 80, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeQuantity);
        Shot('t', model.export_information.quantity_exported, 38, 137.2, 'l');

        /*         let resizeSignature = resizeText(model.);
         */
        let resizeSignature = resizeText(`${model?.signature?.name ?? ''} ${model?.signature?.tax_id ?? ''} ${model?.signature?.date ?? ''}`.trim(), 82, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeSignature);
        Shot('t', String(`( ${`${model?.signature?.name ?? ''} ${model?.signature?.tax_id ?? ''} ${model?.signature?.date ?? ''}`.trim()} )`), 102, 228, 'l');

        let resizeSubmitdate = resizeText(model.submission_date, 38, BODY_SIZE);
        Shot('sf', 'tahoma', 'normal', resizeSubmitdate);
        Shot('t', model.submission_date, 103, 235, 'l');



        const buffer = await Shot("save", 'ลองเปลี่ยนดู');
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});

router.post('/return-cancel/full', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;
        console.log(model);

        // ===== 1) หา page ที่มีข้อมูลจริง =====
        const detail = model.detail || {};
        const pageKeysWithData = Object.keys(detail)
            .filter((name) => {
                const page = detail[name];
                const containers = page?.container || [];
                return containers.some(c => c && Object.keys(c).length > 0);
            })
            .sort();

        const totalPages = pageKeysWithData.length || 1;

        const form_background = await getBase64Image('./app/form_report/return-cancelfull.jpg');
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },
                { name: 'tahoma', vfsName: 'tahoma Normal.ttf', data: stripDataUrl(tahomaNormal), style: 'normal' },
            ]
        });


        const { Shot, CutPadding, state, splitText, resizeText, addPage } = pdf;

        const HEADER_SIZE = 16;
        const BODY_SIZE = 14;
        const ROW = 7;
        const T = (value) => (value ?? "").toString();

        const checkImg = await getBase64Image('./app/form_report/cancelmark.png');
        const drawCheck = (flag, x, y) => {
            if (flag === 'Y') {
                Shot("i", checkImg, x, y, 4.3, 4.3);
            }
        };

        //ลูปวาดทีละหน้า 
        pageKeysWithData.forEach((pageKey, index) => {
            const page = detail[pageKey] || {};
            const containers = page.container || [];

            if (index > 0) {
                addPage();
            }

            Shot("i", form_background, 0, 0, 210, 297);

            const currentPage = index + 1;

            // เลขหน้า
            Shot('sf', 'Cordianew', 'normal', HEADER_SIZE);
            Shot('t', `${currentPage}/${totalPages}`, 178, 20.5, 'l');

            // ชื่อบริษัท
            let resizeExportname = resizeText(model.exporter.exporter_name, 169, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeExportname);
            Shot('t', model.exporter.exporter_name, 35, 27.5, 'l');

            //เลขประจำตัวผู้เสียภาษีอากร
            let resizeTaxid = resizeText(model.exporter.tax_id, 54, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeTaxid);
            Shot('t', model.exporter.tax_id, 52, 35.5, 'l');

            //ที่อยู่
            let resizeAddress = resizeText(
                model.exporter.street_and_number + ' ' +
                model.exporter.sub_province + ' ' +
                model.exporter.district + ' ' +
                model.exporter.provice + ' ' +
                model.exporter.postcode,
                175,
                BODY_SIZE
            );
            Shot('sf', 'CordiaNew', 'normal', resizeAddress);
            Shot('t',
                model.exporter.street_and_number + ' ' +
                model.exporter.sub_province + ' ' +
                model.exporter.district + ' ' +
                model.exporter.provice + ' ' +
                model.exporter.postcode,
                25, 43, 'l'
            );

            //เลขใบขนสินค้าขาออก
            let resizeDecnum = resizeText(model.declaration_no, 29, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeDecnum);
            Shot('t', model.declaration_no, 77, 51.5, 'l');

            // ช่องติ๊กชนิด คอนเทนเนอร์
            drawCheck(model?.container.indicator_container, 19.3, 55.2);

            let resizeContype = resizeText(model.container.container_type, 32, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeContype);
            Shot('t', model.container.container_type, 54, 59.5, 'l');

            let y_car = 63.3;
            drawCheck(model?.car.indicator_car, 19.3, y_car);
            drawCheck(model?.car.indicator_truck, 39.4, y_car);
            drawCheck(model?.car.indicator_closed_truck, 69.5, y_car);
            drawCheck(model?.car.indicator_van, 96.4, y_car);
            drawCheck(model?.car.indicator_pickup, 114.2, y_car);
            drawCheck(model?.car.indicator_tank, 134.5, y_car);
            drawCheck(model?.car.indicator_other, 154.5, y_car);

            let resizeOthercar = resizeText(model.car.other_car, 25, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeOthercar);
            Shot('t', model.car.other_car, 175, 66, 'l');

            let resizeLoadport = resizeText(
                model.load_port.port_name + ' ' + model.load_port.port_code,
                93,
                BODY_SIZE
            );
            Shot('sf', 'CordiaNew', 'normal', resizeLoadport);
            Shot('t',
                model.load_port.port_name + ' ' + '(' + model.load_port.port_code + ')',
                40, 75.5, 'l'
            );

            let resizeVessel = resizeText(model.vessel_name, 38, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeVessel);
            Shot('t', model.vessel_name, 163, 75.5, 'l');

            let resizeVoyagel = resizeText(model.voyagel, 25, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeVoyagel);
            Shot('t', model.voyagel, 29.5, 83.5, 'l');

            let resizeNote = resizeText(model.note, 100, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeNote);
            Shot('t', model.note, 100, 83.5, 'l');


            const col_No = 15;
            const col_transNo = 21;
            const col_contNo = 63;
            const col_conWg = 130;
            const col_conQt = 165;
            const col_conPk = 197;

            let y = 110;

            for (const item of containers) {
                if (!item || Object.keys(item).length === 0) continue;

                Shot('sf', 'Cordianew', 'normal', BODY_SIZE);
                Shot("t", T(item.line_number), col_No, y, 'l');

                const invText = T(item.good_transition_no);
                const invSize = resizeText(invText, 40, BODY_SIZE);
                Shot('sf', 'Cordianew', 'normal', invSize);
                Shot("t", invText, col_transNo, y, 'l');

                const contText = T(item.container_no);
                const contSize = resizeText(contText, 34, BODY_SIZE);
                Shot('sf', 'Cordianew', 'normal', contSize);
                Shot("t", contText, col_contNo, y, 'l');

                const weightText = `${T(item.weight?.weight_value)} ${T(item.weight?.weight_unit)}`.trim();
                const weightSize = resizeText(weightText, 30, BODY_SIZE);
                Shot('sf', 'Cordianew', 'normal', weightSize);
                Shot("t", weightText, col_conWg, y, 'r');

                const quantityText = `${T(item.quantity?.quantity_value)} ${T(item.quantity?.quantity_unit)}`.trim();
                const quantitySize = resizeText(quantityText, 30, BODY_SIZE);
                Shot('sf', 'Cordianew', 'normal', quantitySize);
                Shot("t", quantityText, col_conQt, y, 'r');

                const packText = `${T(item.packing?.pack_value)} ${T(item.packing?.pack_unit)}`.trim();
                const packSize = resizeText(packText, 30, BODY_SIZE);
                Shot('sf', 'Cordianew', 'normal', packSize);
                Shot("t", packText, col_conPk, y, 'r');

                y += ROW;
            }

            //รวม
            const sum = page.summary || {};
            //รวมน้ำหนัก
            if (sum.total_weight) {
                const sumWgText = `${T(sum.total_weight.weight_value)} ${T(sum.total_weight.weight_unit)}`.trim();
                let resizesumWg = resizeText(sumWgText, 30, BODY_SIZE);
                Shot('sf', 'Cordianew', 'normal', resizesumWg);
                Shot("t", sumWgText, 130.5, 252, 'r');
            }
            //รวมปริมาณ
            if (sum.total_quantity) {
                const sumQnText = `${T(sum.total_quantity.quantity_value)} ${T(sum.total_quantity.quantity_unit)}`.trim();
                let resizesumQn = resizeText(sumQnText, 30, BODY_SIZE);
                Shot('sf', 'Cordianew', 'normal', resizesumQn);
                Shot("t", sumQnText, 165, 252, 'r');
            }
            //รวมจำนวนหีบ/ห่อ
            if (sum.total_packing) {
                const sumPkText = `${T(sum.total_packing.pack_value)} ${T(sum.total_packing.pack_unit)}`.trim();
                let resizesumPk = resizeText(sumPkText, 30, BODY_SIZE);
                Shot('sf', 'Cordianew', 'normal', resizesumPk);
                Shot("t", sumPkText, 197, 252, 'r');
            }
        });


        const buffer = await Shot("save", 'return-cancel-full');
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }]);
    }
});

router.post('/exportDeclaration', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;



        let allPage = 0;
        // ใช้งาน:
        const form_background1 = await getBase64Image('./app/form_report/exportDecalration1.jpg');
        const form_background2 = await getBase64Image('./app/form_report/exportDecalration2.jpg');
        const form_background3 = await getBase64Image('./app/form_report/exportDecalration3.jpg');

        // console.log('base64',base64);

        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },

            ]
        });
        const { Shot, CutPadding, state, splitText, resizeText, addPage, setPage } = pdf;
        Shot('sf', 'CordiaNew', 'normal', 10);
        Shot("i", form_background1, 0, 0, 210, 297);
        allPage++;
        Shot("i", model.barcode_base64, 165, 4, 35, 10); // barcode

        state.y = 24.5;
        Shot("fs", 14);
        Shot("t", model?.reference_number || '', 130, state.y, 'l');// ประเภทใบขน
        state.y = 29;
        Shot("t", model?.declaration_number || '', 160, state.y, 'l');// เลขที่ใบขน





        state.y = 28;
        Shot("fs", 14);
        Shot("t", model?.exporter?.tax_number || '', 59, state.y, 'l'); // เลขประจำตัวผู้เสียภาษีอากร
        Shot("t", model?.exporter?.branch ? 'สาขา ' + model?.exporter?.branch || '' : '', 85, state.y, 'l'); // สาขา

        Shot("fs", 9);
        state.y = 32;
        Shot("t", model?.exporter?.thai_name || '', 10, state.y, 'l');
        state.y = 35;
        Shot("t", model?.exporter?.english_name || '', 10, state.y, 'l');
        state.y = 38;
        Shot("t", model?.exporter?.street_and_number || '', 10, state.y, 'l');
        state.y = 41;
        Shot("t", (model?.exporter?.sub_province || '') + ' ' + (model?.exporter?.district || ''), 10, state.y, 'l');
        state.y = 44;
        Shot("t", (model?.exporter?.provice || '') + ' ' + (model?.exporter?.postcode || ''), 10, state.y, 'l');
        if (model?.exporter?.phon_number) {
            state.y = 47;
            Shot("t", 'Tel. ' + (model?.exporter?.phon_number || ''), 10, state.y, 'l');
        }




        Shot("fs", 9);
        state.y = 30;
        Shot("t", model?.is_privilege ? 'ใช้สิทธิประโยชน์' : 'ไม่ใช้สิทธิประโยชน์', 115, state.y, 'l'); // ใช้สิทธิประโยชน์
        let txt_merge = '';
        let index_ = 0;
        for (let item of model?.tax_incentive) {
            if (index_ + 1 == model?.tax_incentive.length) {
                txt_merge += item;
            } else {
                txt_merge += item + ', ';
            }
            index_++;
        }
        state.y = 33;
        Shot("t", txt_merge, 115, state.y, 'l'); // สิทธิประโยชน์
        state.y = 36;
        Shot("t", model?.house_bill_of_lading || '', 115, state.y, 'l');
        state.y = 39;
        Shot("t", model?.master_bill_of_lading || '', 115, state.y, 'l');


        if (model?.invoice_no.length > 0) {
            state.y = 42;
            Shot("t", 'บัญชีราคาสินค้า', 115, state.y, 'l');
        }
        let txt_merge1 = '';
        let index_1 = 0;
        for (let item of model?.invoice_no) {
            if (index_1 + 1 == model?.invoice_no.length) {
                txt_merge1 += item;
            } else {
                txt_merge1 += item + ', ';
            }
            index_1++;
        }
        state.y = 45;

        let cutText_ = splitText(txt_merge1, 90);

        for (let item2 of cutText_) {

            if (item2) {
                Shot("t", item2, 115, state.y, 'l');
                state.y += 3;
            }

        }

        // ชื่อและเลขที่บัตรผ่านพิธีการ
        state.y = 56.15;
        Shot("t", model?.authorised_person?.customs_clearance_id_card || '', 38, state.y, 'l');
        state.y = 59.5;
        Shot("t", model?.authorised_person?.customsclearancename || '', 11, state.y, 'l');

        // ตัวแทนออกของ
        state.y = 68;
        Shot("t", (model?.agent?.tax_number || '') + ' สาขา ' + (model?.agent?.branch || ''), 11, state.y, 'l');
        state.y = 71;
        Shot("t", model?.agent?.thai_name || '', 11, state.y, 'l');

        //สั่งการตรวจ

        Shot("fs", 9);
        let str = '';
        let font_size = 9;
        let max_line = 5;
        let new_font = max_line / model?.time_line_detail.length * font_size;

        state.y = 59.5;
        if (new_font > 9) {
            new_font = 9;
        }
        let line_height = new_font / 3;
        // console.log('new_font', new_font);

        Shot("fs", new_font);
        for (let item of model?.time_line_detail) {
            str += (item?.date_time || '') + ' ' + (item?.description || '') + '\n';
        }
        let cutText = splitText(str, CutPadding('w', 38));

        //  console.log(cutText); 


        for (let item2 of cutText) {
            Shot("t", item2, 115, state.y, 'l');
            state.y += line_height;
        }

        Shot("fs", 9);
        //ชื่อยานพาหนะ
        state.y = 84.15;
        Shot("t", model?.border_transport_means?.vessel_name || '', 28, state.y, 'l');
        state.y = 85;
        Shot("t", model?.duty?.export?.duty_amount || '', CutPadding('w', 73), state.y, 'r');//อากรขาออก
        Shot("t", model?.duty?.export?.duty_amount || '', CutPadding('w', 90.5), state.y, 'r');//ค่าภาษีอากร
        Shot("t", model?.duty?.export?.deposit_amount || '', CutPadding('w', 107), state.y, 'r');//เงินประกัน

        state.y = 95;
        Shot("t", model?.border_transport_means?.mode_code || '', CutPadding('w', -5.25), state.y, 'l');//ส่งออกโดยทาง

        Shot("t", model?.border_transport_means?.departure_date || '', CutPadding('w', 24.25), state.y, 'l');//วันที่ส่งออก

        // เลขที่ชำระภาษีอากร/ประกัน
        Shot("t", model?.duty_exemption || '', 115, state.y, 'l');
        state.y = 90;
        Shot("t", 'วิธีการชำระเงิน ' + model?.bank_info?.payment_method || '', CutPadding('w', 107), state.y, 'r');
        state.y = 95;
        Shot("t", 'วางประกัน : Method = ' + (model?.bank_info?.guarantee_method || '') + ' / Type = ' + (model?.bank_info?.guarantee_type || ''), CutPadding('w', 107), state.y, 'r');

        //ท่าเรือที่ตรวจปล่อยของ
        state.y = 105;
        Shot("t", model?.release_port?.release_port_name || '', CutPadding('w', -5.25), state.y, 'l');
        Shot("t", model?.release_port?.release_port_code || '', CutPadding('w', 22.5), state.y, 'r');

        //ท่่าหรือที่รับบรรทุกของ
        Shot("t", model?.load_port?.load_port_name || '', CutPadding('w', 24), state.y, 'l');
        Shot("t", model?.load_port?.load_port_code || '', CutPadding('w', 53.5), state.y, 'r');

        //ขายไปยังประเทศ
        Shot("t", model?.purchase_country?.purchase_country_name || '', CutPadding('w', 56), state.y, 'l');
        Shot("t", model?.purchase_country?.purchase_country_code || '', CutPadding('w', 78), state.y, 'r');

        //ประเทศปลายทาง
        Shot("t", model?.destination_country?.destination_country_name || '', CutPadding('w', 80), state.y, 'l');
        Shot("t", model?.destination_country?.destination_country_code || '', CutPadding('w', 106.5), state.y, 'r');


        //จำนวนหีบห่อ
        state.y = 118;
        Shot("t", (model?.total_package?.amount || '') + ' ' + (model?.total_package?.amount_unit || '') + '  (' + (model?.total_package?.amount_text || '') + ' ' + (model?.total_package?.amount_unit_text || '') + ')', CutPadding('w', -5.25), state.y, 'l');

        //อัตราแลกเปลี่ยน
        Shot("t", '1 ' + (model?.currency_exchange?.currency || '') + ' ' + (model?.currency_exchange?.exchang_rate || ''), CutPadding('w', 56), state.y, 'l');



        //page1
        let data_last_page = {};
        let permit_over = [];
        let indexPage = 0;
        for (let key of Object.keys(model.detail)) {
            //console.log('Object.keys(model.detail).length-1',Object.keys(model.detail).length-1);
            // แผ่นที่ง



            if (indexPage == 0) {

                data_last_page = model?.detail[key];
                //ใบแรก
                // หน้าแรก ใส่ได้ 3 item
                state.y = 250;
                Shot("rt", model?.send_acc || '', 8, state.y, 'l', 90); //คนส่ง

                state.y = 142;
                for (let item of model?.detail[key].items) {
                    Shot("t", item?.item_number || '', 12, state.y, 'c'); //รายการที่


                    Shot("t", 'Origin : ' + item?.productifo?.origin_country || '', 111, state.y + 2, 'r');
                    Shot("t", 'Pur.County : ' + item?.productifo?.purchase_country || '', 111, state.y + 5, 'r');

                    Shot("t", (item?.productifo?.brand?.brand_name || '') + ' ' + (item?.productifo?.brand?.product_year || ''), 69, state.y + 2, 'l');
                    let reSize_eng = resizeText(item?.productifo?.good_description?.english || '', 80, 9);
                    Shot("fs", reSize_eng);
                    Shot("t", item?.productifo?.good_description?.english || '', 18, state.y + 5, 'l'); //ชื่ออังกฤษ
                    let reSize_th = resizeText(item?.productifo?.good_description?.thai || '', 90, 9);
                    Shot("fs", reSize_th);
                    Shot("t", item?.productifo?.good_description?.thai || '', 18, state.y + 8, 'l'); //ชื่อไทย
                    let reSize_at1 = resizeText(item?.productifo?.good_description?.product_attribute1 || '', 100, 9);
                    Shot("fs", reSize_at1);
                    Shot("t", item?.productifo?.good_description?.product_attribute1 || '', 18, state.y + 11, 'l'); //ลักษณะ1
                    let reSize_at2 = resizeText(item?.productifo?.good_description?.product_attribute2 || '', 100, 9);
                    Shot("fs", reSize_at2);
                    Shot("t", item?.productifo?.good_description?.product_attribute2 || '', 18, state.y + 14, 'l'); //ลักษณะ2
                    Shot("t", item?.productifo?.good_description?.detail_invoice_no ? 'Inv.no.' + (item?.productifo?.good_description?.detail_invoice_no || '') : '', 18, state.y + 17, 'l'); //ลักษณะ invoice

                    Shot("fs", 9);
                    //permit
                    let index_permit = 0;
                    let permit_list = [];
                    for (let item_permit of item?.productifo?.permit) {
                        if (index_permit < 2) {
                            Shot("t", (item_permit.permit_no || '') + ' ' + (item_permit.issue_date || '') + ' ' + (item_permit.tax_id || ''), 111, state.y + 20 + (index_permit * 3), 'r');

                        } else {
                            permit_list.push(item_permit);
                        }
                        index_permit++;
                    }
                    permit_over.push({
                        item_number: item.item_number,
                        permit_list: permit_list

                    });

                    //เครื่องหมายหีบห่อ
                    let cutText_ship_mark = splitText(item?.detail_packakge?.shipping_marks_detail || '', CutPadding('w', 24.5));
                    let index_cutText_ship_mark = 0;
                    for (let item2 of cutText_ship_mark) {
                        Shot("t", item2, 18, state.y - 15 + (3 * index_cutText_ship_mark), 'l');
                        /* state.y += 3; */
                        index_cutText_ship_mark++;
                    }

                    //จำนวนลักษณะและหีบห่อ
                    Shot("t", item?.detail_packakge?.package_amount, 97.5, state.y - 12, 'c');
                    Shot("t", item?.detail_packakge?.package_unit, 97.5, state.y - 9, 'c');

                    //น้ำหลักสุทธิ
                    Shot("t", item?.net_weight?.net_weight_value, CutPadding('w', 70.5), state.y - 15, 'r');
                    Shot("t", item?.net_weight?.unit_net_weight, CutPadding('w', 70.5), state.y - 12, 'r');

                    //ปริมาณ
                    Shot("t", item?.quantity?.quantity, CutPadding('w', 70.5), state.y - 5, 'r');
                    Shot("t", item?.quantity?.UnitCode, CutPadding('w', 70.5), state.y - 2, 'r');

                    //ประเภทพิกัด
                    Shot("t", item?.export_tariff?.export_tariff_code, CutPadding('w', 63), state.y + 5, 'c');

                    //รหัสสถิติหน่วย
                    Shot("t", item?.statitic?.tariff_code, CutPadding('w', 63), state.y + 11.5, 'c');
                    Shot("t", item?.statitic?.statitic_code + ' / ' + item?.statitic?.statitic_unit, CutPadding('w', 63), state.y + 14, 'c');

                    //ราคาของ FOB (ต่างประเทศ)
                    Shot("t", item?.price_foreign?.price_foreign_unit + ' ' + item?.price_foreign?.price_foreign, CutPadding('w', 87.5), state.y - 11, 'r');

                    //ราคาของ FOB (บาท)
                    Shot("t", item?.price_baht, CutPadding('w', 87.5), state.y + 1, 'r');

                    //ราคาประเมินอากร
                    Shot("t", item?.value_assess, CutPadding('w', 87.5), state.y + 13, 'r');

                    //ใช้สิทธิประโยชน์     
                    let txt_tid = '';
                    let index_txd = 0;
                    for (let item_tid of item?.tax_incentive_detail) {
                        if (index_txd + 1 == item?.tax_incentive_detail.length) {
                            txt_tid += item_tid;
                        } else {
                            txt_tid += item_tid + ', ';
                        }

                        index_txd++;
                    }
                    let cutText_tid = splitText(txt_tid, CutPadding('w', 8));
                    let index_cut_tid = 0;
                    for (let item2 of cutText_tid) {
                        Shot("t", item2, CutPadding('w', 89), state.y - 15 + (3 * index_cut_tid), 'l');

                        index_cut_tid++;
                    }

                    //อัตราอากร
                    Shot("t", item?.duty_rate, CutPadding('w', 97.75), state.y + 4.5, 'c');
                    //อากรขาออก
                    Shot("t", item?.duty_amount, CutPadding('w', 106.5), state.y + 13, 'r');

                    //รายละเอียดอื่นๆ
                    Shot("t", item?.other_detail?.nature_of_transacton, CutPadding('w', 56), state.y + 18, 'l');
                    let reSizetxt_rd = resizeText(item?.other_detail?.remark_detapil, 90, 9);
                    Shot('fs', reSizetxt_rd);
                    Shot("t", item?.other_detail?.remark_detapil, CutPadding('w', 56), state.y + 20.5, 'l');
                    Shot("t", item?.other_detail?.reference_declaration_item, CutPadding('w', 106.5), state.y + 18, 'r');
                    let txt_vv = '';
                    let index_vv = 0;
                    for (let item_vv of item?.other_detail?.various_values) {
                        if (index_vv + 1 == item?.other_detail?.various_values.length) {
                            txt_vv += item_vv;
                        } else {
                            txt_vv += item_vv + ', ';
                        }
                        index_vv++;
                    }

                    let reSizetxt_vv = resizeText(txt_vv, 90, 8);
                    Shot('fs', reSizetxt_vv);
                    Shot("t", txt_vv, CutPadding('w', 106.5), state.y + 23, 'r');
                    Shot('fs', 9);



                    state.y += 45;


                }

                //summary 
                state.y = 261.5;

                Shot("t", model?.detail[key]?.summary?.price_foreign?.price_foreign_unit + ' ' + model?.detail[key]?.summary?.price_foreign?.price_foreign_amount, CutPadding('w', 87.5), state.y, 'r');
                state.y = 268.5;
                Shot("t", model?.detail[key]?.summary?.price_amount_baht, CutPadding('w', 87.5), state.y, 'r');

                Shot("t", model?.detail[key]?.summary?.duty_amount, CutPadding('w', 106.5), state.y, 'r');
                state.y = 275.5;
                Shot("t", model?.detail[key]?.summary?.total_duty_amount, CutPadding('w', 106.5), state.y, 'r');

                //footer
                state.y = 280.5;
                Shot("t", model?.detail[key]?.footer?.incoterm, CutPadding('w', 106.5), state.y, 'r');
                state.y = 259;
                Shot("t", model?.detail[key]?.footer?.status, CutPadding('w', -6), state.y, 'l');
                state.y = 262;
                Shot("t", model?.detail[key]?.footer?.status_date, CutPadding('w', -6), state.y, 'l');
                state.y = 259;
                Shot("t", 'Total G.W. :' + model?.detail[key]?.footer?.total_gross_weight?.gross_weight_value + ' ' + model?.detail[key]?.footer?.total_gross_weight?.gross_weight_unit, CutPadding('w', 15), state.y, 'l');
                state.y = 262;
                Shot("t", 'Total QTY. :' + model?.detail[key]?.footer?.total_quality?.quantity_value + ' ' + model?.detail[key]?.footer?.total_quality?.quality_unit, CutPadding('w', 15), state.y, 'l');

                state.y = 259;
                Shot("t", 'Total N.W. :' + model?.detail[key]?.footer?.total_net_weight?.net_weight_value + ' ' + model?.detail[key]?.footer?.total_net_weight?.net_weight_unit, CutPadding('w', 45), state.y, 'l');
                state.y = 262;
                Shot("t", 'Total QTY. :' + model?.detail[key]?.footer?.total_invoice_quality?.quantity_value + ' ' + model?.detail[key]?.footer?.total_invoice_quality?.quality_unit, CutPadding('w', 45), state.y, 'l');
                state.y = 265;
                Shot("t", 'Total PACK' + model?.detail[key]?.footer?.total_pack?.pack_value + ' ' + model?.detail[key]?.footer?.total_pack?.pack_unit, CutPadding('w', 15), state.y, 'l');

                state.y = 278;
                Shot("t", '(' + (model?.detail[key]?.footer?.signature?.signature_name || '') + ')', CutPadding('w', 51), state.y, 'l');
                state.y = 280.5;
                Shot("t", (model?.detail[key]?.footer?.signature?.signature_id || ''), CutPadding('w', 51), state.y, 'l');
                state.y = 283.5;
                Shot("t", (model?.detail[key]?.footer?.signature_date || ''), CutPadding('w', 51), state.y, 'l');




            }

            if (indexPage > 0 && indexPage < Object.keys(model?.detail).length - 2) {
                allPage++;
                data_last_page = model?.detail[key];
                //ใบต่อ ที่ n
                addPage();
                Shot("i", form_background2, 0, 0, 210, 297);
                state.y = 250;
                Shot("rt", model?.send_acc || '', 8, state.y, 'l', 90); //คนส่ง

                state.y = 15;
                Shot("fs", 14);
                Shot("t", model?.reference_number || '', CutPadding('w', 75), state.y, 'l');// ประเภทใบขน
                state.y = 19;
                Shot("t", model?.declaration_number || '', CutPadding('w', 75), state.y, 'l');// เลขที่ใบขน
                Shot("fs", 9);
                state.y = 51;
                for (let item of model?.detail[key].items) {
                    Shot("t", item?.item_number || '', 12, state.y, 'c'); //รายการที่


                    Shot("t", 'Origin : ' + item?.productifo?.origin_country || '', 111, state.y + 2, 'r');
                    Shot("t", 'Pur.County : ' + item?.productifo?.purchase_country || '', 111, state.y + 5, 'r');

                    Shot("t", (item?.productifo?.brand?.brand_name || '') + ' ' + (item?.productifo?.brand?.product_year || ''), 69, state.y + 2, 'l');
                    

                    Shot("t", (item?.productifo?.brand?.brand_name || '') + ' ' + (item?.productifo?.brand?.product_year || ''), 69, state.y + 2, 'l');
                    let reSize_eng = resizeText(item?.productifo?.good_description?.english || '', 80, 9);
                    Shot("fs", reSize_eng);
                    Shot("t", item?.productifo?.good_description?.english || '', 18, state.y + 5, 'l'); //ชื่ออังกฤษ
                    let reSize_th = resizeText(item?.productifo?.good_description?.thai || '', 90, 9);
                    Shot("fs", reSize_th);
                    Shot("t", item?.productifo?.good_description?.thai || '', 18, state.y + 8, 'l'); //ชื่อไทย
                    let reSize_at1 = resizeText(item?.productifo?.good_description?.product_attribute1 || '', 100, 9);
                    Shot("fs", reSize_at1);
                    Shot("t", item?.productifo?.good_description?.product_attribute1 || '', 18, state.y + 11, 'l'); //ลักษณะ1
                    let reSize_at2 = resizeText(item?.productifo?.good_description?.product_attribute2 || '', 100, 9);
                    Shot("fs", reSize_at2);
                    Shot("t", item?.productifo?.good_description?.product_attribute2 || '', 18, state.y + 14, 'l'); //ลักษณะ2
                    Shot("t", item?.productifo?.good_description?.detail_invoice_no ? 'Inv.no.' + (item?.productifo?.good_description?.detail_invoice_no || '') : '', 18, state.y + 17, 'l'); //ลักษณะ invoice

                    Shot("fs", 9);
                  /*   Shot("t", item?.productifo?.good_description?.english || '', 18, state.y + 5, 'l'); //ชื่ออังกฤษ
                    Shot("t", item?.productifo?.good_description?.thai || '', 18, state.y + 8, 'l'); //ชื่อไทย
                    Shot("t", item?.productifo?.good_description?.product_attribute1 || '', 18, state.y + 11, 'l'); //ลักษณะ1
                    Shot("t", item?.productifo?.good_description?.product_attribute2 || '', 18, state.y + 14, 'l'); //ลักษณะ2
                    Shot("t", item?.productifo?.good_description?.detail_invoice_no ? 'Inv.no.' + (item?.productifo?.good_description?.detail_invoice_no || '') : '', 18, state.y + 17, 'l'); //ลักษณะ invoice */
                    //permit
                    let index_permit = 0;
                    let permit_list = [];
                    for (let item_permit of item?.productifo?.permit) {
                        if (index_permit < 2) {
                            Shot("t", (item_permit.permit_no || '') + ' ' + (item_permit.issue_date || '') + ' ' + (item_permit.tax_id || ''), 111, state.y + 20 + (index_permit * 3), 'r');

                        } else {
                            permit_list.push(item_permit);
                        }
                        index_permit++;
                    }
                    permit_over.push({
                        item_number: item.item_number,
                        permit_list: permit_list

                    });

                    //เครื่องหมายหีบห่อ
                    let cutText_ship_mark = splitText(item?.detail_packakge?.shipping_marks_detail || '', CutPadding('w', 24.5));
                    let index_cutText_ship_mark = 0;
                    for (let item2 of cutText_ship_mark) {
                        Shot("t", item2, 18, state.y - 15 + (3 * index_cutText_ship_mark), 'l');
                        /* state.y += 3; */
                        index_cutText_ship_mark++;
                    }

                    //จำนวนลักษณะและหีบห่อ
                    Shot("t", item?.detail_packakge?.package_amount, 97.5, state.y - 12, 'c');
                    Shot("t", item?.detail_packakge?.package_unit, 97.5, state.y - 9, 'c');

                    //น้ำหลักสุทธิ
                    Shot("t", item?.net_weight?.net_weight_value, CutPadding('w', 70.5), state.y - 15, 'r');
                    Shot("t", item?.net_weight?.unit_net_weight, CutPadding('w', 70.5), state.y - 12, 'r');

                    //ปริมาณ
                    Shot("t", item?.quantity?.quantity, CutPadding('w', 70.5), state.y - 5, 'r');
                    Shot("t", item?.quantity?.UnitCode, CutPadding('w', 70.5), state.y - 2, 'r');

                    //ประเภทพิกัด
                    Shot("t", item?.export_tariff?.export_tariff_code, CutPadding('w', 63), state.y + 5, 'c');

                    //รหัสสถิติหน่วย
                    Shot("t", item?.statitic?.tariff_code, CutPadding('w', 63), state.y + 11.5, 'c');
                    Shot("t", item?.statitic?.statitic_code + ' / ' + item?.statitic?.statitic_unit, CutPadding('w', 63), state.y + 14, 'c');

                    //ราคาของ FOB (ต่างประเทศ)
                    Shot("t", item?.price_foreign?.price_foreign_unit + ' ' + item?.price_foreign?.price_foreign, CutPadding('w', 87.5), state.y - 11, 'r');

                    //ราคาของ FOB (บาท)
                    Shot("t", item?.price_baht, CutPadding('w', 87.5), state.y + 1, 'r');

                    //ราคาประเมินอากร
                    Shot("t", item?.value_assess, CutPadding('w', 87.5), state.y + 13, 'r');

                    //ใช้สิทธิประโยชน์     
                    let txt_tid = '';
                    let index_txd = 0;
                    for (let item_tid of item?.tax_incentive_detail) {
                        if (index_txd + 1 == item?.tax_incentive_detail.length) {
                            txt_tid += item_tid;
                        } else {
                            txt_tid += item_tid + ', ';
                        }

                        index_txd++;
                    }
                    let cutText_tid = splitText(txt_tid, CutPadding('w', 8));
                    let index_cut_tid = 0;
                    for (let item2 of cutText_tid) {
                        Shot("t", item2, CutPadding('w', 89), state.y - 15 + (3 * index_cut_tid), 'l');

                        index_cut_tid++;
                    }

                    //อัตราอากร
                    Shot("t", item?.duty_rate, CutPadding('w', 97.75), state.y + 4.5, 'c');
                    //อากรขาออก
                    Shot("t", item?.duty_amount, CutPadding('w', 106.5), state.y + 13, 'r');

                    //รายละเอียดอื่นๆ
                    Shot("t", item?.other_detail?.nature_of_transacton, CutPadding('w', 56), state.y + 18, 'l');
                    let reSizetxt_rd = resizeText(item?.other_detail?.remark_detapil, 90, 9);
                    Shot('fs', reSizetxt_rd);
                    Shot("t", item?.other_detail?.remark_detapil, CutPadding('w', 56), state.y + 20.5, 'l');
                    Shot("t", item?.other_detail?.reference_declaration_item, CutPadding('w', 106.5), state.y + 18, 'r');
                    let txt_vv = '';
                    let index_vv = 0;
                    for (let item_vv of item?.other_detail?.various_values) {
                        if (index_vv + 1 == item?.other_detail?.various_values.length) {
                            txt_vv += item_vv;
                        } else {
                            txt_vv += item_vv + ', ';
                        }
                        index_vv++;
                    }

                    let reSizetxt_vv = resizeText(txt_vv, 90, 9);
                    Shot('fs', reSizetxt_vv);
                    Shot("t", txt_vv, CutPadding('w', 106.5), state.y + 23, 'r');
                    Shot('fs', 9);



                    state.y += 45;


                }

                //summary 
                state.y = 261;

                Shot("t", model?.detail[key]?.summary?.price_foreign?.price_foreign_unit + ' ' + model?.detail[key]?.summary?.price_foreign?.price_foreign_amount, CutPadding('w', 87.5), state.y, 'r');
                state.y = 268;
                Shot("t", model?.detail[key]?.summary?.price_amount_baht, CutPadding('w', 87.5), state.y, 'r');

                Shot("t", model?.detail[key]?.summary?.duty_amount, CutPadding('w', 106.5), state.y, 'r');
                state.y = 275;
                Shot("t", model?.detail[key]?.summary?.total_duty_amount, CutPadding('w', 106.5), state.y, 'r');

                //footer
                state.y = 279;
                Shot("t", model?.detail[key]?.footer?.incoterm, CutPadding('w', 106.5), state.y, 'r');
                state.y = 258;
                Shot("t", model?.detail[key]?.footer?.status, CutPadding('w', -6), state.y, 'l');
                state.y = 261;
                Shot("t", model?.detail[key]?.footer?.status_date, CutPadding('w', -6), state.y, 'l');
                state.y = 258;
                Shot("t", 'Total G.W. :' + model?.detail[key]?.footer?.total_gross_weight?.gross_weight_value + ' ' + model?.detail[key]?.footer?.total_gross_weight?.gross_weight_unit, CutPadding('w', 15), state.y, 'l');
                state.y = 261;
                Shot("t", 'Total QTY. :' + model?.detail[key]?.footer?.total_quality?.quantity_value + ' ' + model?.detail[key]?.footer?.total_quality?.quality_unit, CutPadding('w', 15), state.y, 'l');

                state.y = 258;
                Shot("t", 'Total N.W. :' + model?.detail[key]?.footer?.total_net_weight?.net_weight_value + ' ' + model?.detail[key]?.footer?.total_net_weight?.net_weight_unit, CutPadding('w', 45), state.y, 'l');
                state.y = 261;
                Shot("t", 'Total QTY. :' + model?.detail[key]?.footer?.total_invoice_quality?.quantity_value + ' ' + model?.detail[key]?.footer?.total_invoice_quality?.quality_unit, CutPadding('w', 45), state.y, 'l');
                state.y = 264;
                Shot("t", 'Total PACK' + model?.detail[key]?.footer?.total_pack?.pack_value + ' ' + model?.detail[key]?.footer?.total_pack?.pack_unit, CutPadding('w', 15), state.y, 'l');

                state.y = 276;
                Shot("t", '(' + (model?.detail[key]?.footer?.signature?.signature_name || '') + ')', CutPadding('w', 51), state.y, 'l');
                state.y = 279.5;
                Shot("t", (model?.detail[key]?.footer?.signature?.signature_id || ''), CutPadding('w', 51), state.y, 'l');
                state.y = 283.5;
                Shot("t", (model?.detail[key]?.footer?.signature_date || ''), CutPadding('w', 51), state.y, 'l');
            }


            indexPage++;
        }



        if (model.detail?.page_shipping_mark?.shipping_mark.length > 0) {
            addPage();
            Shot("i", form_background2, 0, 0, 210, 297);
            allPage++;

            state.y = 250;
            Shot("rt", model?.send_acc || '', 8, state.y, 'l', 90); //คนส่ง

            state.y = 15;
            Shot("fs", 14);
            Shot("t", model?.reference_number || '', CutPadding('w', 75), state.y, 'l');// ประเภทใบขน
            state.y = 19;
            Shot("t", model?.declaration_number || '', CutPadding('w', 75), state.y, 'l');// เลขที่ใบขน
            Shot("fs", 9);


            //summary 
            state.y = 261;

            Shot("t", data_last_page?.summary?.price_foreign?.price_foreign_unit + ' ' + data_last_page?.summary?.price_foreign?.price_foreign_amount, CutPadding('w', 87.5), state.y, 'r');
            state.y = 268;
            Shot("t", data_last_page?.summary?.price_amount_baht, CutPadding('w', 87.5), state.y, 'r');

            Shot("t", data_last_page?.summary?.duty_amount, CutPadding('w', 106.5), state.y, 'r');
            state.y = 275;
            Shot("t", data_last_page?.summary?.total_duty_amount, CutPadding('w', 106.5), state.y, 'r');

            //footer
            state.y = 279;
            Shot("t", data_last_page?.footer?.incoterm, CutPadding('w', 106.5), state.y, 'r');
            state.y = 258;
            Shot("t", data_last_page?.footer?.status, CutPadding('w', -6), state.y, 'l');
            state.y = 261;
            Shot("t", data_last_page?.footer?.status_date, CutPadding('w', -6), state.y, 'l');
            state.y = 258;
            Shot("t", 'Total G.W. :' + data_last_page?.footer?.total_gross_weight?.gross_weight_value + ' ' + data_last_page?.footer?.total_gross_weight?.gross_weight_unit, CutPadding('w', 15), state.y, 'l');
            state.y = 261;
            Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_quality?.quantity_value + ' ' + data_last_page?.footer?.total_quality?.quality_unit, CutPadding('w', 15), state.y, 'l');

            state.y = 258;
            Shot("t", 'Total N.W. :' + data_last_page?.footer?.total_net_weight?.net_weight_value + ' ' + data_last_page?.footer?.total_net_weight?.net_weight_unit, CutPadding('w', 45), state.y, 'l');
            state.y = 261;
            Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_invoice_quality?.quantity_value + ' ' + data_last_page?.footer?.total_invoice_quality?.quality_unit, CutPadding('w', 45), state.y, 'l');
            state.y = 264;
            Shot("t", 'Total PACK' + data_last_page?.footer?.total_pack?.pack_value + ' ' + data_last_page?.footer?.total_pack?.pack_unit, CutPadding('w', 15), state.y, 'l');

            state.y = 276;
            Shot("t", '(' + (data_last_page?.footer?.signature?.signature_name || '') + ')', CutPadding('w', 51), state.y, 'l');
            state.y = 279.5;
            Shot("t", (data_last_page?.footer?.signature?.signature_id || ''), CutPadding('w', 51), state.y, 'l');
            state.y = 283.5;
            Shot("t", (data_last_page?.footer?.signature_date || ''), CutPadding('w', 51), state.y, 'l');


            state.y = 72;

            Shot("t", 'Shipping Mark', CutPadding('w', 0), state.y, 'l');
            state.y += 4;
            let count_sh = 1;
            for (let item_sh of model.detail?.page_shipping_mark?.shipping_mark) {
                if (count_sh > 60) {
                    addPage();
                    Shot("i", form_background2, 0, 0, 210, 297);
                    allPage++;
                    state.y = 250;
                    Shot("rt", model?.send_acc || '', 8, state.y, 'l', 90); //คนส่ง

                    state.y = 15;
                    Shot("fs", 14);
                    Shot("t", model?.reference_number || '', CutPadding('w', 75), state.y, 'l');// ประเภทใบขน
                    state.y = 19;
                    Shot("t", model?.declaration_number || '', CutPadding('w', 75), state.y, 'l');// เลขที่ใบขน
                    Shot("fs", 9);
                    count_sh = 1;

                    //summary 
                    state.y = 261;

                    Shot("t", data_last_page?.summary?.price_foreign?.price_foreign_unit + ' ' + data_last_page?.summary?.price_foreign?.price_foreign_amount, CutPadding('w', 87.5), state.y, 'r');
                    state.y = 268;
                    Shot("t", data_last_page?.summary?.price_amount_baht, CutPadding('w', 87.5), state.y, 'r');

                    Shot("t", data_last_page?.summary?.duty_amount, CutPadding('w', 106.5), state.y, 'r');
                    state.y = 275;
                    Shot("t", data_last_page?.summary?.total_duty_amount, CutPadding('w', 106.5), state.y, 'r');

                    //footer
                    state.y = 279;
                    Shot("t", data_last_page?.footer?.incoterm, CutPadding('w', 106.5), state.y, 'r');
                    state.y = 258;
                    Shot("t", data_last_page?.footer?.status, CutPadding('w', -6), state.y, 'l');
                    state.y = 261;
                    Shot("t", data_last_page?.footer?.status_date, CutPadding('w', -6), state.y, 'l');
                    state.y = 258;
                    Shot("t", 'Total G.W. :' + data_last_page?.footer?.total_gross_weight?.gross_weight_value + ' ' + data_last_page?.footer?.total_gross_weight?.gross_weight_unit, CutPadding('w', 15), state.y, 'l');
                    state.y = 261;
                    Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_quality?.quantity_value + ' ' + data_last_page?.footer?.total_quality?.quality_unit, CutPadding('w', 15), state.y, 'l');

                    state.y = 258;
                    Shot("t", 'Total N.W. :' + data_last_page?.footer?.total_net_weight?.net_weight_value + ' ' + data_last_page?.footer?.total_net_weight?.net_weight_unit, CutPadding('w', 45), state.y, 'l');
                    state.y = 261;
                    Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_invoice_quality?.quantity_value + ' ' + data_last_page?.footer?.total_invoice_quality?.quality_unit, CutPadding('w', 45), state.y, 'l');
                    state.y = 264;
                    Shot("t", 'Total PACK' + data_last_page?.footer?.total_pack?.pack_value + ' ' + data_last_page?.footer?.total_pack?.pack_unit, CutPadding('w', 15), state.y, 'l');

                    state.y = 276;
                    Shot("t", '(' + (data_last_page?.footer?.signature?.signature_name || '') + ')', CutPadding('w', 51), state.y, 'l');
                    state.y = 279.5;
                    Shot("t", (data_last_page?.footer?.signature?.signature_id || ''), CutPadding('w', 51), state.y, 'l');
                    state.y = 283.5;
                    Shot("t", (data_last_page?.footer?.signature_date || ''), CutPadding('w', 51), state.y, 'l');

                    state.y = 72;
                    Shot("t", 'Shipping Mark', CutPadding('w', 0), state.y, 'l');
                    state.y += 4;


                }
                Shot("t", item_sh, CutPadding('w', 0), state.y, 'l');
                state.y += 3;

                count_sh++;
            }
            /*   Shot("t", model.detail?.page_shipping_mark?.shipping_mark, CutPadding('w', 0), state.y, 'l'); */








        }

        if (model.detail?.page_invoice?.invoice && model.detail.page_invoice.invoice.length > 0) {
            addPage();
            Shot("i", form_background2, 0, 0, 210, 297);
            allPage++;
            state.y = 10.5;
            /*      Shot("t", indexPage + ' / ' + (pageKeys.length + 1), CutPadding('w', 96.5), state.y, 'l');  */
            let txt_invoice = '';
            let index_invoice = 0;

            for (let item of (model.detail?.page_invoice?.invoice || [])) {
                if (index_invoice == model.detail?.page_invoice?.invoice.length - 1) {
                    txt_invoice += item;
                } else {
                    txt_invoice += item + ', ';
                }
                index_invoice++;
            }
            let cutTextInvoice = splitText(txt_invoice, CutPadding('w', 100));
            state.y = 72;

            Shot("t", 'Invoice No. :', CutPadding('w', 0), state.y, 'l');
            for (let item2 of cutTextInvoice) {
                Shot("t", item2, CutPadding('w', 7.5), state.y, 'l');
                state.y += 3;
            }

            //summary ใช้ data_last_page
            state.y = 237.5;
            Shot("t", (data_last_page?.summary?.price_amount || data_last_page?.summary?.price_foreign?.price_foreign_amount || ''), CutPadding('w', 30), state.y, 'r'); //รวม ราคาของ (เงินต่างประเทศ)
            state.y = 248.5;
            Shot("t", (data_last_page?.summary?.price_amount_baht || ''), CutPadding('w', 30), state.y, 'r'); //รวม ราคาของ (บาท)
            state.y = 242;
            Shot("fs", 16);
            Shot("t", (data_last_page?.summary?.duty_amount || ''), CutPadding('w', 44.5), state.y, 'r'); // รวม ค่าอากรขาเข้า
            state.y = 245;
            Shot("fs", 10);
            Shot("t", (data_last_page?.summary?.quantity_amount || data_last_page?.summary?.total_quality?.quantity_value || ''), CutPadding('w', 44.5), state.y, 'r'); // รวม ปริมาณ
            state.y = 241;
            Shot("fs", 16);
            Shot("t", (data_last_page?.summary?.duty_fee || ''), CutPadding('w', 59), state.y, 'r'); // รวม ค่าธรรมเนียม
            state.y = 252.5;
            Shot("t", (data_last_page?.summary?.other_duty || ''), CutPadding('w', 59), state.y, 'r'); // รวม ภาษีอื่นๆ

            state.y = 241;
            Shot("t", (data_last_page?.summary?.excise_duty || ''), CutPadding('w', 87), state.y, 'r'); // รวม สรรพสามิต
            state.y = 252.5;
            Shot("t", (data_last_page?.summary?.interior_duty || ''), CutPadding('w', 87), state.y, 'r'); // รวม ภาษีมหาดไทย
            state.y = 245.5;
            Shot("t", (data_last_page?.summary?.vat_duty || ''), CutPadding('w', 102), state.y, 'r'); // รวม ภาษีมูลค่าเพิ่ม

            state.y = 258.5;
            Shot("t", (data_last_page?.summary?.total_duty_amount || ''), CutPadding('w', 102), state.y, 'r'); // รวมค่าภาษีอากรทั้งสิ้น

            Shot("rt", model?.send_acc || '', 8, state.y, 'l', 90); //คนส่ง

            state.y = 15;
            Shot("fs", 14);
            Shot("t", model?.reference_number || '', CutPadding('w', 75), state.y, 'l');// ประเภทใบขน
            state.y = 19;
            Shot("t", model?.declaration_number || '', CutPadding('w', 75), state.y, 'l');// เลขที่ใบขน
            Shot("fs", 9);


            //summary 
            state.y = 261;

            Shot("t", data_last_page?.summary?.price_foreign?.price_foreign_unit + ' ' + data_last_page?.summary?.price_foreign?.price_foreign_amount, CutPadding('w', 87.5), state.y, 'r');
            state.y = 268;
            Shot("t", data_last_page?.summary?.price_amount_baht, CutPadding('w', 87.5), state.y, 'r');

            Shot("t", data_last_page?.summary?.duty_amount, CutPadding('w', 106.5), state.y, 'r');
            state.y = 275;
            Shot("t", data_last_page?.summary?.total_duty_amount, CutPadding('w', 106.5), state.y, 'r');

            //footer
            state.y = 279;
            Shot("t", data_last_page?.footer?.incoterm, CutPadding('w', 106.5), state.y, 'r');
            state.y = 258;
            Shot("t", data_last_page?.footer?.status, CutPadding('w', -6), state.y, 'l');
            state.y = 261;
            Shot("t", data_last_page?.footer?.status_date, CutPadding('w', -6), state.y, 'l');
            state.y = 258;
            Shot("t", 'Total G.W. :' + data_last_page?.footer?.total_gross_weight?.gross_weight_value + ' ' + data_last_page?.footer?.total_gross_weight?.gross_weight_unit, CutPadding('w', 15), state.y, 'l');
            state.y = 261;
            Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_quality?.quantity_value + ' ' + data_last_page?.footer?.total_quality?.quality_unit, CutPadding('w', 15), state.y, 'l');

            state.y = 258;
            Shot("t", 'Total N.W. :' + data_last_page?.footer?.total_net_weight?.net_weight_value + ' ' + data_last_page?.footer?.total_net_weight?.net_weight_unit, CutPadding('w', 45), state.y, 'l');
            state.y = 261;
            Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_invoice_quality?.quantity_value + ' ' + data_last_page?.footer?.total_invoice_quality?.quality_unit, CutPadding('w', 45), state.y, 'l');
            state.y = 264;
            Shot("t", 'Total PACK' + data_last_page?.footer?.total_pack?.pack_value + ' ' + data_last_page?.footer?.total_pack?.pack_unit, CutPadding('w', 15), state.y, 'l');

            state.y = 276;
            Shot("t", '(' + (data_last_page?.footer?.signature?.signature_name || '') + ')', CutPadding('w', 51), state.y, 'l');
            state.y = 279.5;
            Shot("t", (data_last_page?.footer?.signature?.signature_id || ''), CutPadding('w', 51), state.y, 'l');
            state.y = 283.5;
            Shot("t", (data_last_page?.footer?.signature_date || ''), CutPadding('w', 51), state.y, 'l');
        }

        if (permit_over.length > 0) {
            addPage();
            Shot("i", form_background2, 0, 0, 210, 297);
            allPage++;

            Shot("rt", model?.send_acc || '', 8, state.y, 'l', 90); //คนส่ง

            state.y = 15;
            Shot("fs", 14);
            Shot("t", model?.reference_number || '', CutPadding('w', 75), state.y, 'l');// ประเภทใบขน
            state.y = 19;
            Shot("t", model?.declaration_number || '', CutPadding('w', 75), state.y, 'l');// เลขที่ใบขน
            Shot("fs", 9);


            //summary 
            state.y = 261;

            Shot("t", data_last_page?.summary?.price_foreign?.price_foreign_unit + ' ' + data_last_page?.summary?.price_foreign?.price_foreign_amount, CutPadding('w', 87.5), state.y, 'r');
            state.y = 268;
            Shot("t", data_last_page?.summary?.price_amount_baht, CutPadding('w', 87.5), state.y, 'r');

            Shot("t", data_last_page?.summary?.duty_amount, CutPadding('w', 106.5), state.y, 'r');
            state.y = 275;
            Shot("t", data_last_page?.summary?.total_duty_amount, CutPadding('w', 106.5), state.y, 'r');

            //footer
            state.y = 279;
            Shot("t", data_last_page?.footer?.incoterm, CutPadding('w', 106.5), state.y, 'r');
            state.y = 258;
            Shot("t", data_last_page?.footer?.status, CutPadding('w', -6), state.y, 'l');
            state.y = 261;
            Shot("t", data_last_page?.footer?.status_date, CutPadding('w', -6), state.y, 'l');
            state.y = 258;
            Shot("t", 'Total G.W. :' + data_last_page?.footer?.total_gross_weight?.gross_weight_value + ' ' + data_last_page?.footer?.total_gross_weight?.gross_weight_unit, CutPadding('w', 15), state.y, 'l');
            state.y = 261;
            Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_quality?.quantity_value + ' ' + data_last_page?.footer?.total_quality?.quality_unit, CutPadding('w', 15), state.y, 'l');

            state.y = 258;
            Shot("t", 'Total N.W. :' + data_last_page?.footer?.total_net_weight?.net_weight_value + ' ' + data_last_page?.footer?.total_net_weight?.net_weight_unit, CutPadding('w', 45), state.y, 'l');
            state.y = 261;
            Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_invoice_quality?.quantity_value + ' ' + data_last_page?.footer?.total_invoice_quality?.quality_unit, CutPadding('w', 45), state.y, 'l');
            state.y = 264;
            Shot("t", 'Total PACK' + data_last_page?.footer?.total_pack?.pack_value + ' ' + data_last_page?.footer?.total_pack?.pack_unit, CutPadding('w', 15), state.y, 'l');

            state.y = 276;
            Shot("t", '(' + (data_last_page?.footer?.signature?.signature_name || '') + ')', CutPadding('w', 51), state.y, 'l');
            state.y = 279.5;
            Shot("t", (data_last_page?.footer?.signature?.signature_id || ''), CutPadding('w', 51), state.y, 'l');
            state.y = 283.5;
            Shot("t", (data_last_page?.footer?.signature_date || ''), CutPadding('w', 51), state.y, 'l');


            state.y = 72;

            Shot("t", 'Permit', CutPadding('w', 0), state.y, 'l');
            state.y += 4;
            let count_ph = 1;
            for (let item_ph of permit_over) {

                Shot("t", 'รายการที่ : ' + item_ph.item_number, CutPadding('w', 0), state.y, 'l');

                for (let item_detail_permit of item_ph.permit_list) {
                    if (count_ph > 60) {
                        addPage();
                        Shot("i", form_background2, 0, 0, 210, 297);
                        allPage++;
                        state.y = 250;
                        Shot("rt", model?.send_acc || '', 8, state.y, 'l', 90); //คนส่ง

                        state.y = 15;
                        Shot("fs", 14);
                        Shot("t", model?.reference_number || '', CutPadding('w', 75), state.y, 'l');// ประเภทใบขน
                        state.y = 19;
                        Shot("t", model?.declaration_number || '', CutPadding('w', 75), state.y, 'l');// เลขที่ใบขน
                        Shot("fs", 9);
                        count_ph = 1;

                        //summary 
                        state.y = 261;

                        Shot("t", data_last_page?.summary?.price_foreign?.price_foreign_unit + ' ' + data_last_page?.summary?.price_foreign?.price_foreign_amount, CutPadding('w', 87.5), state.y, 'r');
                        state.y = 268;
                        Shot("t", data_last_page?.summary?.price_amount_baht, CutPadding('w', 87.5), state.y, 'r');

                        Shot("t", data_last_page?.summary?.duty_amount, CutPadding('w', 106.5), state.y, 'r');
                        state.y = 275;
                        Shot("t", data_last_page?.summary?.total_duty_amount, CutPadding('w', 106.5), state.y, 'r');

                        //footer
                        state.y = 279;
                        Shot("t", data_last_page?.footer?.incoterm, CutPadding('w', 106.5), state.y, 'r');
                        state.y = 258;
                        Shot("t", data_last_page?.footer?.status, CutPadding('w', -6), state.y, 'l');
                        state.y = 261;
                        Shot("t", data_last_page?.footer?.status_date, CutPadding('w', -6), state.y, 'l');
                        state.y = 258;
                        Shot("t", 'Total G.W. :' + data_last_page?.footer?.total_gross_weight?.gross_weight_value + ' ' + data_last_page?.footer?.total_gross_weight?.gross_weight_unit, CutPadding('w', 15), state.y, 'l');
                        state.y = 261;
                        Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_quality?.quantity_value + ' ' + data_last_page?.footer?.total_quality?.quality_unit, CutPadding('w', 15), state.y, 'l');

                        state.y = 258;
                        Shot("t", 'Total N.W. :' + data_last_page?.footer?.total_net_weight?.net_weight_value + ' ' + data_last_page?.footer?.total_net_weight?.net_weight_unit, CutPadding('w', 45), state.y, 'l');
                        state.y = 261;
                        Shot("t", 'Total QTY. :' + data_last_page?.footer?.total_invoice_quality?.quantity_value + ' ' + data_last_page?.footer?.total_invoice_quality?.quality_unit, CutPadding('w', 45), state.y, 'l');
                        state.y = 264;
                        Shot("t", 'Total PACK' + data_last_page?.footer?.total_pack?.pack_value + ' ' + data_last_page?.footer?.total_pack?.pack_unit, CutPadding('w', 15), state.y, 'l');

                        state.y = 276;
                        Shot("t", '(' + (data_last_page?.footer?.signature?.signature_name || '') + ')', CutPadding('w', 51), state.y, 'l');
                        state.y = 279.5;
                        Shot("t", (data_last_page?.footer?.signature?.signature_id || ''), CutPadding('w', 51), state.y, 'l');
                        state.y = 283.5;
                        Shot("t", (data_last_page?.footer?.signature_date || ''), CutPadding('w', 51), state.y, 'l');

                        state.y = 72;
                        Shot("t", 'Permit', CutPadding('w', 0), state.y, 'l');
                        state.y += 4;


                    }
                    Shot("t", (item_detail_permit.permit_no || '') + ' ' + (item_detail_permit.issue_date || '') + ' ' + (item_detail_permit.tax_id || ''), CutPadding('w', 8), state.y, 'l');
                    state.y += 3;
                    count_ph++;
                }


            }
            /*   Shot("t", model.detail?.page_shipping_mark?.shipping_mark, CutPadding('w', 0), state.y, 'l'); */








        }


        //ใบปิด
        addPage();

        Shot("i", form_background3, 0, 0, 210, 297);
        allPage++;
        Shot("fs", 14);
        state.y = 17;
        Shot("t", model?.declaration_number || '', CutPadding('w', 64), state.y, 'l');
        Shot("fs", 12);
        state.y = 28;

        if (model?.endorse?.inspection_request) {
            Shot("t", 'Inspection Request', CutPadding('w', -3), state.y, 'l');
            Shot("t", model?.endorse?.inspection_request || '', CutPadding('w', 15), state.y, 'l');
            state.y += 3;
        }

        if (model?.endorse?.reassessment_request) {
            Shot("t", 'Assessment Request', CutPadding('w', -3), state.y, 'l');
            Shot("t", model?.endorse?.reassessment_request || '', CutPadding('w', 15), state.y, 'l');
            state.y += 3;
        }

        if (model?.endorse?.cargo_packing_type) {
            Shot("t", 'Cargo Packing Type', CutPadding('w', -3), state.y, 'l');
            Shot("t", model?.endorse?.cargo_packing_type || '', CutPadding('w', 15), state.y, 'l');
            state.y += 3;
        }

        Shot("fs", 10);
        state.y = 115;
        let index_con = 0;

        let font_size_cd = 10;
        let max_line_cd = 40;
        let new_font_cd = max_line_cd / model?.endorse?.cotainer_detail.length * font_size_cd;


        if (new_font_cd > 10) {
            new_font_cd = 10;
        }
        let line_height_cd = new_font_cd / 3.5;
        // console.log('new_font', new_font);

        Shot("fs", new_font_cd);
        for (let item of model?.endorse?.cotainer_detail) {
            Shot("t", (index_con + 1) + '. ' + (item.container_no || ''), CutPadding('w', -3), state.y, 'l');
            Shot("t", (item.vessel || '') + ' / ' + (item.voyagel || ''), CutPadding('w', 43), state.y, 'l');

            Shot("t", (item.release_port || '') + ' / ' + (item.load_port || ''), CutPadding('w', 80), state.y, 'l');


            state.y += line_height_cd;
            index_con++;
        }
        // console.log('state.y',state.y);

        /*  state.y = 235; */
        /*    Shot("t", '---------------------------------------', CutPadding('w', -3) , state.y,'l'); */
        let start_y = 117.85;

        Shot("r", CutPadding('w', 0), state.y, CutPadding('w', 90.5), 145 - (state.y - start_y));
        Shot("fs", 10);

        state.y += 4;
        Shot("t", 'โอนสิทธิ BOI และอื่นๆ', CutPadding('w', 1), state.y, 'l');
        state.y += 4;
        let cutTextNote = splitText(model?.endorse?.note || '', CutPadding('w', 89));
        for (let item of cutTextNote) {

            Shot("t", item, CutPadding('w', 1), state.y, 'l');
            state.y += 3
        }


        state.y = 267;
        Shot("t", model?.endorse?.footer?.status + '   ' + model?.endorse?.footer?.status_date, CutPadding('w', 0), state.y, 'l');
        Shot("t", model?.endorse?.footer?.status + '   ' + model?.endorse?.footer?.status_date, CutPadding('w', 0), state.y, 'l');
        Shot("t", model?.endorse?.footer?.status + '   ' + model?.endorse?.footer?.status_date, CutPadding('w', 0), state.y, 'l');

        state.y = 274;
        Shot("t", '(' + (model?.endorse?.footer?.signature?.signature_name || '') + ')', CutPadding('w', 57), state.y, 'l');
        state.y = 277;
        Shot("t", (model?.endorse?.footer?.signature?.signature_id || ''), CutPadding('w', 57), state.y, 'l');
        state.y = 281.5;
        Shot("t", (model?.endorse?.footer?.submission_date || ''), CutPadding('w', 57), state.y, 'l');


        /*        console.log(allPage); */

        for (let i = 0; i < allPage; i++) {
            setPage(i + 1);
            //เลขหน้า
            Shot("fs", 9);
            if (i == 0) {
                state.y = 19;
                Shot("t", (i + 1) + ' / ' + allPage, 180, state.y, 'l');
            }
            if (i > 0 && i + 1 < allPage) {
                state.y = 17;
                Shot("t", (i + 1) + ' / ' + allPage, 187, state.y, 'l');
            }

            if (i + 1 == allPage) {
                state.y = 16;
                Shot("t", (i + 1) + ' / ' + allPage, 189, state.y, 'l');// เลขที่ใบขน
            }

        }




        //console.log('permit_over',permit_over);

        const filename = req?.query?.file_name || "ใบขนสินค้าขาเข้า"; // ชื่อไฟล์เป็นภาษาไทย


        const buffer = await Shot("save", filename);
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.setHeader("Content-Type", "application/pdf");

        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});

router.post('/return-cancel/partial', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;
        console.log(model);
        const detail = model.detail || {};
        const pageKeysWithData = Object.keys(detail)
            .filter((name) => {
                const page = detail[name];
                const containers = page?.container || [];
                return containers.some(c => c && Object.keys(c).length > 0);
            })
            .sort();

        const totalPages = pageKeysWithData.length || 1;


        const form_background = await getBase64Image('./app/form_report/return-cancelpartial.jpg');
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            orientation: "landscape",
            fonts: [
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },
                { name: 'tahoma', vfsName: 'tahoma Normal.ttf', data: stripDataUrl(tahomaNormal), style: 'normal' },
            ]
        });


        const { Shot, CutPadding, state, splitText, resizeText, addPage } = pdf;

        const checkImg = await getBase64Image('./app/form_report/cancelmark.png');
        const HEADER_SIZE = 16;
        const BODY_SIZE = 14;
        const ROW = 10;

        const T = (value) => (value ?? "").toString();

        const drawCheck = (flag, x, y) => {
            if (flag === 'Y') {
                Shot("i", checkImg, x, y, 4.3, 4.3);
            }
        };

        //ลูปวาดทีละหน้า 
        pageKeysWithData.forEach((pageKey, index) => {
            const page = detail[pageKey] || {};
            const containers = page.container || [];

            if (index > 0) {
                addPage();
            }

            Shot("i", form_background, 0, 0, 297, 210);

            const currentPage = index + 1;

            // เลขหน้า
            Shot('sf', 'CordiaNew', 'normal', HEADER_SIZE);
            Shot('t', `${currentPage}/${totalPages}`, 267, 20.5, 'l');

            // ชื่อบริษัท
            let resizeExportname = resizeText(model.exporter.exporter_name, 160, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeExportname);
            Shot('t', model.exporter.exporter_name, 35, 27.5, 'l');

            //เลขประจำตัวผู้เสียภาษีอากร
            let resizeTaxid = resizeText(model.exporter.tax_id, 35, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeTaxid);
            Shot('t', model.exporter.tax_id, 240, 27.5, 'l');

            //ที่อยู่
            let resizeAddress = resizeText(
                model.exporter.street_and_number + ' ' +
                model.exporter.sub_province + ' ' +
                model.exporter.district + ' ' +
                model.exporter.provice + ' ' +
                model.exporter.postcode,
                260,
                BODY_SIZE
            );
            Shot('sf', 'CordiaNew', 'normal', resizeAddress);
            Shot('t',
                model.exporter.street_and_number + ' ' +
                model.exporter.sub_province + ' ' +
                model.exporter.district + ' ' +
                model.exporter.provice + ' ' +
                model.exporter.postcode,
                22, 35.5, 'l'
            );

            //เลขใบขนสินค้าขาออก
            let resizeDecnum = resizeText(model.declaration_no, 29, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeDecnum);
            Shot('t', model.declaration_no, 77, 43.6, 'l');

            // ช่องติ๊กชนิด คอนเทนเนอร์
            drawCheck(model?.container.indicator_container, 19.3, 47.5);

            let resizeContype = resizeText(model.container.container_type, 32, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeContype);
            Shot('t', model.container.container_type, 54, 51.5, 'l');

            let y_car = 55.5;
            drawCheck(model?.indicator_car, 19.3, y_car);
            drawCheck(model?.indicator_truck, 39.4, y_car);
            drawCheck(model?.indicator_closed_truck, 69.5, y_car);
            drawCheck(model?.indicator_van, 96.4, y_car);
            drawCheck(model?.indicator_pickup, 114.2, y_car);
            drawCheck(model?.indicator_tank, 134.5, y_car);
            drawCheck(model?.indicator_other, 154.5, y_car);

            let resizeOthercar = resizeText(model.other_car, 25, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeOthercar);
            Shot('t', model.other_car, 175, 58.5, 'l');


            //ส่งออกท่า
            let resizeLoadport = resizeText(model.load_port.port_name + ' ' + model.load_port.port_code, 93, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeLoadport);
            Shot('t',
                model.load_port.port_name + ' ' + '(' + model.load_port.port_code + ')',
                40, 67.5, 'l'
            );

            //โดยยานพาหนะ
            let resizeVessel = resizeText(model.vessel_name, 50, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeVessel);
            Shot('t', model.vessel_name, 163, 67.5, 'l');

            let resizeVoyagel = resizeText(model.voyagel, 25, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeVoyagel);
            Shot('t', model.voyagel, 246, 67.5, 'l');

            let resizeNote = resizeText(model.note, 100, BODY_SIZE);
            Shot('sf', 'CordiaNew', 'normal', resizeNote);
            Shot('t', model.note, 117, 75.5, 'l');

            // ===== ตารางรายละเอียดของ "หน้านี้เท่านั้น" =====
            const col_No = 15;     //ลำดับ
            const col_transNo = 21;     //เลขที่ใบกำกับการขนย้ายสินค้า
            const col_contNo = 63;     //หมายเลขตู้คอนเทนเนอร์
            const col_typeDec = 98;     //ชนิดของรายการที่ในใบขน
            const col_conWg = 214;    //น้ำหนัก
            const col_conQt = 237;    //ปริมาณ
            const col_conPk = 259;    //จำนวนหีบ/ห่อ
            const col_conBaht = 288;    //ราคาFOBบาท

            let y = 103;

            for (const item of containers) {
                if (!item || Object.keys(item).length === 0) continue;

                Shot('sf', 'CordiaNew', 'normal', BODY_SIZE);
                Shot("t", T(item.line_number), col_No, y, 'l');

                //เลขที่ใบกำกับการขนย้ายสินค้า
                const invText = T(item.good_transition_no);
                const invSize = resizeText(invText, 40, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', invSize);
                Shot("t", invText, col_transNo, y, 'l');

                //หมายเลขตู้คอนเทนเนอร์
                const contText = T(item.container_no);
                const contSize = resizeText(contText, 34, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', contSize);
                Shot("t", contText, col_contNo, y, 'l');

                //ชนิดของรายการที่ในใบขน
                const typeText = `${T(item.line_number_declaration)}`.trim();
                const typeSize = resizeText(typeText, 90, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', typeSize);
                Shot("t", typeText, col_typeDec, y, 'l');

                //น้ำหนัก
                const weightText = `${T(item.weight?.weight_value)} ${T(item.weight?.weight_unit)}`.trim();
                const weightSize = resizeText(weightText, 24, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', weightSize);
                Shot("t", weightText, col_conWg, y, 'r');
                //ปริมาณ
                const quantityText = `${T(item.quantity?.quantity_value)} ${T(item.quantity?.quality_unit)}`.trim();
                const quantitySize = resizeText(quantityText, 22, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', quantitySize);
                Shot("t", quantityText, col_conQt, y, 'r');
                //จำนวนหีบห่อ
                const packText = `${T(item.packing?.pack_value)} ${T(item.packing?.pack_unit)}`.trim();
                const packSize = resizeText(packText, 19, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', packSize);
                Shot("t", packText, col_conPk, y, 'r');
                //ราคาเงินบาท
                const bahtText = `${T(item.price_baht)}`.trim();
                const bahtSize = resizeText(bahtText, 27, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', bahtSize);
                Shot("t", bahtText, col_conBaht, y, 'r');

                y += ROW;
            }

            const sum = page.summary || {};
            if (sum.total_weight) {
                const sumWgText = `${T(sum.total_weight.weight_value)} ${T(sum.total_weight.weight_unit)}`.trim();
                let resizesumWg = resizeText(sumWgText, 24, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', resizesumWg);
                Shot("t", sumWgText, 214, 175, 'r');
            }

            if (sum.total_quantity) {
                const sumQtText = `${T(sum.total_quantity.quantity_value)} ${T(sum.total_quantity.quality_unit)}`.trim();
                let resizesumQt = resizeText(sumQtText, 21, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', resizesumQt);
                Shot("t", sumQtText, 239, 175, 'r');
            }

            if (sum.total_packing) {
                const sumPkText = `${T(sum.total_packing.pack_value)} ${T(sum.total_packing.pack_unit)}`.trim();
                let resizesumPk = resizeText(sumPkText, 19, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', resizesumPk);
                Shot("t", sumPkText, 259, 175, 'r');
            }



            if (sum.total_price_baht) {
                const sumPbText = `${T(sum.total_price_baht)}`.trim();
                let resizesumPb = resizeText(sumPbText, 27, BODY_SIZE);
                Shot('sf', 'CordiaNew', 'normal', resizesumPb);
                Shot("t", sumPbText, 288, 175, 'r');
            }
        });


        const buffer = await Shot("save", 'return-cancel-full');
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }]);
    }
});


router.post('/changeTransportAir', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;

        // ใช้งาน:
        const form_background = await getBase64Image('./app/form_report/changeTransportAir.jpg');


        // console.log('base64',base64);
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                /*    { name: 'THSarabunNew', vfsName: 'THSarabunNew.ttf', data: Get_font_pdf_th2(), style: 'bold' }, */
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },
            ]
        });
        const { Shot, CutPadding, state, resizeText } = pdf;
        Shot("i", form_background, 0, 0, 210, 297);
        /*  Shot("i", model.qrcode_base64, 18, 40, 30, 30);//qr_code
               Shot("i", model.barcode_base64, 55, 42, 130, 18);//bar_code */

        Shot("fs", 10);

        Shot('sf', 'CordiaNew', 'normal', 10);


        state.y = 33;
        Shot("t", model?.exporter_name || '', CutPadding('w', 4), state.y, 'l');

        state.y = 39;
        Shot("t", model?.goods_transition_no || '', CutPadding('w', 30), state.y, 'l');


        state.y = 73;
        Shot("t", model?.old_detail?.vessel_name || '', CutPadding('w', -1), state.y, 'l');
        Shot("t", model?.old_detail?.departure_date || '', CutPadding('w', 10), state.y, 'l');
        Shot("t", (model?.old_detail?.export_areas?.area_code || '') + ' ' + (model?.old_detail?.export_areas?.area_name || ''), CutPadding('w', 60), state.y, 'l');

        Shot("i", model?.old_detail?.master_airway_bill, 60, 62.5, 20, 10);//bar_code
        Shot("t", 'Master', CutPadding('w', 21), 72, 'l');

        Shot("i", model?.old_detail?.house_airway_bill, 60, 74.5, 20, 10);//bar_code
        Shot("t", 'House', CutPadding('w', 21), 84, 'l');


        state.y = 102.5;
        Shot("t", model?.new_detail?.vessel_name || '', CutPadding('w', -1), state.y, 'l');
        Shot("t", model?.new_detail?.departure_date || '', CutPadding('w', 10), state.y, 'l');
        Shot("t", (model?.new_detail?.export_areas?.area_code || '') + ' ' + (model?.new_detail?.export_areas?.area_name || ''), CutPadding('w', 60), state.y, 'l');

        Shot("i", model?.old_detail?.master_airway_bill, 60, 91.5, 20, 10);//bar_code
        Shot("t", 'Master', CutPadding('w', 21), 101, 'l');

        Shot("i", model?.old_detail?.house_airway_bill, 60, 103.5, 20, 10);//bar_code
        Shot("t", 'House', CutPadding('w', 21), 113, 'l');







        const buffer = await Shot("save", 'return-cancel-full');
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }]);
    }
});

router.post('/changeExprotPort', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;

        // ใช้งาน:
        const form_background = await getBase64Image('./app/form_report/changeExportPort.jpg');


        // console.log('base64',base64);
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                /*    { name: 'THSarabunNew', vfsName: 'THSarabunNew.ttf', data: Get_font_pdf_th2(), style: 'bold' }, */
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },
            ]
        });
        const { Shot, CutPadding, state, resizeText } = pdf;
        Shot("i", form_background, 0, 0, 210, 297);
        /*  Shot("i", model.qrcode_base64, 18, 40, 30, 30);//qr_code
               Shot("i", model.barcode_base64, 55, 42, 130, 18);//bar_code */

        Shot("fs", 14);

        Shot('sf', 'CordiaNew', 'normal', 14);
        state.y = 25.5;
        Shot("t", '1/1', CutPadding('w', 98), state.y, 'l');
        state.y = 32.5;
        let resizeTh = resizeText(model?.company_info?.thai_name || '',CutPadding('w', 47),14);
        Shot("fs",resizeTh);
        Shot("t", model?.company_info?.thai_name || '', CutPadding('w', 2), state.y, 'l');
        Shot("fs",14);
        Shot("t", model?.company_info?.tax_id || '', CutPadding('w', 83), state.y, 'l');
        state.y = 38.5;
        let address = (model?.company_info?.street_and_number || '') + ' ' + (model?.company_info?.sub_province || '') + ' ' + (model?.company_info?.district || '') + ' ' + (model?.company_info?.provice || '') + ' ' + (model?.company_info?.postcode || '') + (model?.company_info?.phon_number ? ' Tel. '+model?.company_info?.phon_number : '');
        let resizeAd= resizeText(address,CutPadding('w', 92),14);
        Shot("fs",resizeAd);
        Shot("t", address, CutPadding('w', 0), state.y, 'l');
        Shot("fs",14);
        state.y = 51.5;
        Shot("t", (model?.transport_information?.old_load_port?.port_name || '') + ' (' + (model?.transport_information?.old_load_port?.port_code || '') + ')', CutPadding('w', 10), state.y, 'l');
        state.y = 52;
        Shot("t",'____________________________________________________________________________________________________', CutPadding('w', 10), state.y, 'l');
        state.y = 57.75;
        Shot("t", model?.transport_information?.vessel || '', CutPadding('w', 8.5), state.y, 'l');
        Shot("t", model?.transport_information?.voyage_number || '', CutPadding('w', 86), state.y, 'l');
        state.y = 64;
        Shot("t", (model?.transport_information?.new_load_port?.port_name || '') + ' (' + (model?.transport_information?.new_load_port?.port_code || '') + ')', CutPadding('w', 42), state.y, 'l');

        state.y = 85;

        Shot("t", model?.container_information?.line_no || '', CutPadding('w', -2), state.y, 'c');
        Shot("t", model?.container_information?.goods_transition_number || '', CutPadding('w', 3), state.y, 'l');
        Shot("t", model?.container_information?.container_no || '', CutPadding('w', 30.5), state.y, 'l');
        Shot("t", (model?.container_information?.package_amount?.package_value || '') + ' ' + (model?.container_information?.package_amount?.package_unit || ''), CutPadding('w', 82.5), state.y, 'r');
        Shot("t", (model?.container_information?.gross_weight?.weight_value || '') + ' ' + (model?.container_information?.gross_weight?.weight_unit || ''), CutPadding('w', 104), state.y, 'r');
        /*     state.y += 4; */


        state.y = 257;
        Shot("t", (model?.summary?.package_amount?.package_value || '') + ' ' + (model?.summary?.package_amount?.package_unit || ''), CutPadding('w', 82.5), state.y, 'r');
        Shot("t", (model?.summary?.gross_weight?.weight_value || '') + ' ' + (model?.summary?.gross_weight?.weight_unit || ''), CutPadding('w', 104), state.y, 'r');







        const buffer = await Shot("save", 'return-cancel-full');
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }]);
    }
});

router.post('/incomplete-items', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {


        let model = req.body;
        console.log(model);

        const form_background = await getBase64Image('./app/form_report/incomplete-items.jpg');
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            orientation: "landscape",
            fonts: [
                /*   { name: 'THSarabunNew', vfsName: 'THSarabunNew.ttf', data: Get_font_pdf_th2(), style: 'bold' }, */
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },

            ]
        });
        const { Shot, CutPadding, state, splitText, resizeText } = pdf;
        Shot("i", form_background, 0, 0, 297, 210);

        const HEADER_SIZE = 16;
        const BODY_SIZE = 14;
        const checkImg = await getBase64Image('./app/form_report/cancelmark.png');
        const drawCheck = (flag, x, y) => {
            if (flag === 'Y') {
                Shot("i", checkImg, x, y, 4.3, 4.3);
            }
        };

        Shot('sf', 'CordiaNew', 'normal', 14);

        Shot("t", 'Ref. no. : ' + model.reference_number, 9, 18.5, 'l');  // Ref No.


        let resizeName = resizeText(model.exporter.exporter_name, 50, BODY_SIZE);   //ชื่อบริษัท
        Shot('sf', 'CordiaNew', 'normal', resizeName);
        Shot("t", model.exporter.exporter_name, 33, 46.5, 'l');

        let resizeTaxid = resizeText(model.exporter.tax_id, 50, BODY_SIZE);    //เลขประจำตัวผู้เสียภาษีอากร
        Shot('sf', 'CordiaNew', 'normal', resizeTaxid);
        Shot("t", model.exporter.tax_id, 223, 46.5, 'l');

        const addressText = [model?.exporter?.street_and_number + ' ' +        //ที่อยู่
            model?.exporter?.sub_province + ' ' +
            model?.exporter?.district + ' ' +
            model?.exporter?.provice + ' ' +
            model?.exporter?.postcode]
        let resizeAddress = resizeText(addressText, 100, BODY_SIZE);
        Shot("t", addressText, 23, 53.5, 'l');

        state.y = 63.5;

        drawCheck(model.container.indicator_container, 21.4, state.y); //ช่องติ๊กถูก คอนเทนเนอร์ ชนิด
        let resizeContype = resizeText(model?.container?.container_type, 50, BODY_SIZE);
        Shot('sf', 'CordiaNew', 'normal', resizeContype);
        Shot("t", model?.container?.container_type, 30, state.y, 'l');





        const buffer = await Shot("save", 'ลองเปลี่ยนดู');
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});



router.post('/final_price_detail_170', [/* authenticateToken */], async (req, res, next) => {
    const response = new Responsedata(req, res);
    try {
        let model = req.body;




        // ใช้งาน:
        const form_background = await getBase64Image('./app/form_report/imgmock170_01.png');
        /*        const form_background2 = await getBase64Image('./app/form_report/delivery-transfer-notes_b.jpg'); */
        function textseting(keytext) {
            return keytext ? (keytext).toString() : ''
        }

        // console.log('base64',base64);
        const stripDataUrl = s => (s.startsWith('data:') ? s.split(',')[1] : s);
        const pdf = createPdfTools({
            format: "a4",
            fonts: [
                { name: 'CordiaNew', vfsName: 'CORDIA.ttf', data: stripDataUrl(cordiaNormal), style: 'normal' },
                { name: 'CordiaNew', vfsName: 'Cordia New Bold.ttf', data: stripDataUrl(cordiaBold), style: 'bold' },


            ]
        });
        const { Shot, CutPadding, state, addPage, splitText } = pdf;


        let list = Object.keys(model.detail)
            .filter(key => /^page_\d+$/.test(key))
            .sort((a, b) => Number(a.split("_")[1]) - Number(b.split("_")[1]))
            .map(key => model.detail[key]);
        console.log(list);

        Shot("i", form_background, 0, -5, 210, 297);



        const filename = req?.query?.file_name || "แบบฟอร์มชำระค่าภาษี"; // ชื่อไฟล์เป็นภาษาไทย
        console.log('filename');
        const buffer = await Shot("save", filename);
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
        res.send(buffer);

        /*   return response.success(tempRes);
   */
    } catch (error) {
        return response.error([{
            errorcode: 400,
            errorDis: error.message
        }])
    }
});


module.exports = router;