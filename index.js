const express = require('express')
const { google } = require('googleapis')
const bodyParser = require('body-parser')
const app = express()

app.set('view engine', 'ejs')
app.use(bodyParser.json())
app.use(bodyParser.urlencoded({ extended: true }))
app.use(express.static(__dirname + '/views'))

/************************************************************ กำหนด variableต่างๆ ทีใช้งาน *********************************************************** */
// กำหนดตัวแปรสำหรับใช้งาน mySQL
var mysql = require('mysql2');
var Stock = mysql.createConnection({
    host: "localhost",
    user: "Mulan",
    password: "Mulan*220542",
    database: "login"// ชื่อ Database ที่สร้างไว้
});

// กำหนดตัวแปรสำหรับใส่คำสั่ง MySQL
var get_mysql_data = (sql, place_holder) => {
    return new Promise(function (resolve, reject) { //กำหนดให้ return  object Promise รอ

        Stock.connect(() => {  //รันคำสั่ง SQL
            Stock.query(sql, place_holder, (err, result) => {

                if (result == null) {
                    return reject({ message: "Result is Empty" });
                }
                else if (err) {
                    console.log(err);
                    return reject(err);
                }
                
                resolve(result);  //ส่งผลลัพธืของคำสั่ง sql กลับไปให้ทำงานต่อ
            })

        });

    });
}

//กำหนดตัวแปรสำหรับใช้งาน Google Sheet
const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json", // ไฟล์ที่ใช้ในการระบุตัวตนเพื่อเข้าถึข้อมูล
    scopes: " https://www.googleapis.com/auth/spreadsheets", // สโคปของการใช้งานฐานข้อมูลสามารถเปลี่ยนได้ตามความเหมาะสม

})
const client = auth.getClient();
const googleSheets = google.sheets({ version: "v4", auth: client });
const spreadsheetId = "15IW-DqCfTUYrcylSBpqmGtVay4W2PKkwFWhW8_IyM68"; // รหัสของชีสที่จะเข้าถึง

/*********************************************************** กำหนด function ต่างๆ ทีใช้งาน ***************************************************************** */

// ฟังก์ชั่นสำหรับดึงข้อมูลจาก Google Sheet
async function GetData_Googlesheet(get_range, get_majorDimension) {

    const getData = await googleSheets.spreadsheets.values.get({ //รอเก็บข้อมูลรายชื่อจาก Googlesheet
        auth,
        spreadsheetId,
        range: get_range, // ข้อมูลที่ใส่เป็น String Ex "ชีต1!I2:I",
        majorDimension: get_majorDimension // ข้อมูลที่ใส่เป็น String Ex "ROWS"
    });
    return getData // ส่งคืนค่าที่ได้จาก Google Sheet
}

// ฟังก์ชั่นสำหรับเพิ่มข้อมูลใส่ Google Sheet
async function SentData_Googlesheet(sent_range, sent_values) {

    googleSheets.spreadsheets.values.append({
        auth,
        spreadsheetId,
        range: sent_range,
        valueInputOption: "USER_ENTERED",
        resource: {
            values: [
                sent_values
            ],
        },
    });
}

// ฟังก์ชั่นสำหรับอัพเดตข้อมูลใส่ Google Sheet
async function UpdateData_Googlesheet(sent_range, sent_values, Dimension) {

    googleSheets.spreadsheets.values.update({
        auth,
        spreadsheetId,
        valueInputOption: "USER_ENTERED",
        range: sent_range, //String EX "Warehouse!C2"
        resource: {
            majorDimension: Dimension , // String EX "COLUMNS"
            values: [sent_values] // Array
        }

    });
}

//ฟังก์ชั้นอัพเดตข้อมูล Warehouse & เพิ่มรายการ Stocklist
async function Update_Warehouse_Stocklist_Googlesheet(request,response, Status, Product_Name) {

    var form = request.body
    var Value_Departure = {
        Check: null,
        Admin: null,
        Status: Status, // ใส่เป็น String
        Required_Date: form.Required_Date,
        Product_Name: Product_Name, //ใส่เป็น String
        Model: form.Model,
        Amount: form.Amount,
        Hospital: form.Hospital,
        Operator: form.Operator,
        Note: form.Note
    }

    // ดึงจำนวนคงเหลือตามชื่อสินค้าที่ใส่
    const getProduct_Sheet = await GetData_Googlesheet("Warehouse!A2:E", "ROWS")//เก็บชื่อสินค้าทั้งหมดจาก Warehouse
    const numRows = getProduct_Sheet.data.values.length

    for (let i = 0; i < numRows; i++) {
        var getProduct = Object.values(getProduct_Sheet.data.values)
        var getData = getProduct[i]
        var checkProduct = getData.includes(Value_Departure.Model)

        if (checkProduct == true) {
            var getTotal = getData[3]
            var MaxTotal = getData[4]
            var ProductRow = i // หมายเลข Row ที่มีชื่อสินค้าอยู่
            break;
        }
    }

    //ตรวจสอบสถานะ
    if (Status == "Departure") {
        var Remain = parseInt(getTotal) - parseInt(Value_Departure.Amount)
    }
    else if (Status == "Return") {
        var Remain = parseInt(getTotal) + parseInt(Value_Departure.Amount)
    }

    //ตรวจสอบจำนวนสินค้าคงเหลือ
    if (Remain < 0) { // ของไม่พอ ยืมมากกว่าจำนวนที่มี
        var result_Remain = "สินค้าไม่เพียงพอ คงเหลือ" + getTotal + " ชิ้น กรุณาติดต่อผู้ดูเเลสินค้าหรือ ดำเนินการในภายหลัง"
        var result_Update = false
    }
    else if (Remain > MaxTotal) { // คืนเกินจำนวนที่ควรจะเป็น
        var result_Remain = "ใส่จำนวนสินค้ามากกว่าสินค้าที่มีในคลัง กรุณาดำเนินการใหม่อีกครั้ง"
        var result_Update = false
    }
    else if (0 <= Remain <= MaxTotal) { // ของพอต่อการยืมหรือคืนไม่เกินจำนวนที่ควรจะเป็น
        var result_Remain = "สินค้าคงเหลือ" + Remain.toString() + "ชิ้น"
        var result_Update = true

        // เพิ่มข้อมูล Stocklist ใน GoogleSheet
        var sent_values = Object.values(Value_Departure)
        SentData_Googlesheet("Stock_List!A:J", sent_values)

        // อัพเดตข้อมูลยอดคงเหลือปัจจุบันใน Warehouse 
        var RemainData = [];
        for (let j = 0; j < numRows; j++) {
            if (j == ProductRow) {
                RemainData.push(Remain.toString())
            }
            else {
                RemainData.push(null)
            }
        }
        UpdateData_Googlesheet("Warehouse!D2",  RemainData, "COLUMNS")
    }
    else {
        response.render('login', { result_login: "System Error Please Contact Service" })
        var UpdateResult = "System Error"
    }
 
    //ตรวจสอบสถานะการอัพเดตสินค้าในคลัง
    if (result_Update == true){
        response.render('success', { result_warehouse: result_Remain })
        var UpdateResult = "Success"
    }
    else {
        response.render('fail', { result_warehouse: result_Remain })
         var UpdateResult = "Fail"
    }

    return ([UpdateResult,result_Remain])
}

/************************************* สำหรับจัดการข้อมูล Username & Password ************************************************* */
app.all('/', function (request, response) {
    response.render('login')
})

// Login
app.all('/login', async (request, response) => {

    if (request.method == 'GET') {
        response.render('login')
    }

    else if (request.method == 'POST') {
        var form = request.body
        let Login = form.Username_Login;
        let Password = form.password_login; //รับข้อมูลจากหน้าเว็บไซต์

        var sql1 = "SELECT * FROM userpass WHERE Username = '" + Login + "' AND Password = '" + Password + "'"; //ตรวจสอบข้อมูลว่ามี User กับ Pass ที่ตรงกันหรือไม่
        check_UserPass = await get_mysql_data(sql1)
        var check = check_UserPass.length //  ตรวจสอบจำ Row ถ้า row = 1 มีข้อมูล row = 0 ไม่มีข้อมูล row > 1 คือข้อมูลซ้ำ

        if (Login == "Saowaporn" && Password == "220542") { // สำหรับ Admin ดูข้อมูลการเข้าสู่ระบบ
            response.redirect('UserPass_Info')
        }

        else if (check == 1) { // User & Pass ถูกต้อง
            response.redirect('/departure-return')
            console.log("Suscess")
        }
        else { // User ถูก Pass ผิด
            response.render('login', { result_login: "User or Password Incorrect Please TRY AGIAN" })
            console.log("Pass Error")
        }
    }
})

// ลงทะเบียน
app.all('/signin', async (request, response) => {

    if (request.method == 'GET') {
        response.render('signin')
    }

    else if (request.method == 'POST') {
        let form = request.body
        let data_signin = [
            [
                form.Date,
                form.title,
                form.Firstname,
                form.Lastname,
                form.Gentle,
                form.Phone,
                form.Email,
                form.Username,
                form.Password
            ]
        ]

        if (form.Password == form.Comfirmpassword) { // หากรหัสผ่านถูกต้อง

            var sql = "INSERT INTO userpass (Date, Title, FirstName, LastName, Gentle, PhoneNumber, Email, Username, Password) VALUES ?";
            insert_UserPass = await get_mysql_data(sql, [data_signin])
            response.render('login')
        }
        else {
            response.render('signin')
        }
    }
})

// ยืนยันตัวตนเพื่อเปลี่ยน Password
app.all('/resetpassword', async (request, response) => {

    if (request.method == 'GET') {
        response.render('resetpassword')
    }
    else if (request.method == 'POST') {

        let Email = request.body.Email_Reset

        var sql = "SELECT * FROM userpass WHERE Email = '" + Email + "'";
        check_Email = await get_mysql_data(sql)

        var check_length = check_Email.length // ตรวจสอบว่ามีข้อมูลในฐานข้อมูลหรือไม่ ถ้า row = 1 มี row = 0 ไม่มี row > 1 คือข้อมูลซ้ำ
        const check_ID = check_Email[0].ID

        if (check_length == 1) { // Email ถูกต้อง
            response.redirect('/edit/' + check_ID)
        }
        else {
            response.render('resetpassword', { result_reset: "Email Account Not Found Please TRY AGIAN" })
        }

    }
})

// เเก้ไขข้อมูล User & Password
app.all('/edit/:check_ID', async function (request, response) {

    if (request.method == 'GET') {
        if (request.params.check_ID) {

            var sql = "SELECT * FROM userpass WHERE ID = '" + request.params.check_ID + "'"; // เลือกข้อมูลจาก ID ที่ถูก request 
            var get_Data_ID = await get_mysql_data(sql)
            const sent_Data_ID = get_Data_ID[0] // ดึงข้อมูลจาก Array เพื่อนำไปเเสดงที่หน้า Edit
            response.render('edit', { data: sent_Data_ID })

        } else {
            response.render('login')
        }

    } else if (request.method == 'POST') {

        let form = request.body
        let data_Edit = [
            [
                form.Date,
                form.title,
                form.Firstname,
                form.Lastname,
                form.Gentle,
                form.Phone,
                form.Email,
                form.Username,
                form.Password
            ]
        ]

        if (form.Password == form.Comfirmpassword) { // หากรหัสผ่านถูกต้อง

            var sql1 = "DELETE FROM userpass WHERE Email = '" + form.Email + "'"; // ลบข้อมูลโดยตรวจสอบจาก Email ที่ใส่เข้ามา
            console.log(form.Email)
            var check_Email = await get_mysql_data(sql1)
            var check = check_Email.affectedRows // ตรวจสอบว่ามีการ Delete เกิดขึ้นหรือไม่ ถ้ามี check = 1 ไม่มี check = 0

            if (check == 1) { // หากมี Email จึงจะเพิ่มข้อมูล
                var sql = "INSERT INTO userpass (Date, Title, FirstName, LastName, Gentle, PhoneNumber, Email, Username, Password) VALUES ?";
                insert_UserPass_Edit = await get_mysql_data(sql, [data_Edit])
                response.redirect('/login')
            }
            else {
                console.log('Email not Found')
                var sql = "SELECT * FROM userpass WHERE ID = '" + request.params.check_ID + "'"; // เลือกข้อมูลจาก ID ที่ถูก request 
                var get_Data_ID = await get_mysql_data(sql)
                const sent_Data_ID = get_Data_ID[0] // ดึงข้อมูลจาก Array เพื่อนำไปเเสดงที่หน้า Edit
                response.render('edit', { data: sent_Data_ID, result_Edit: "The Email were not found in the database Please TRY AGAIN." })

            }
        }

        else {// หากใส่ Password เเละ ConfirmPassword ไม่ตรงก็อยู่หน้าเดิม
            console.log('Pass & Comfirm Fail')
            var sql = "SELECT * FROM userpass WHERE ID = '" + request.params.check_ID + "'"; // เลือกข้อมูลจาก ID ที่ถูก request 
            var get_Data_ID = await get_mysql_data(sql)
            const sent_Data_ID = get_Data_ID[0] // ดึงข้อมูลจาก Array เพื่อนำไปเเสดงที่หน้า Edit
            response.render('edit', { data: sent_Data_ID })

        }
    }
})

//เสดงข้อมูล User & Password
app.all('/UserPass_Info', async function (request, response) {

    var get_data = await get_mysql_data("SELECT * FROM userpass")
    response.render('UserPass_Info', { get_data })

})

//ลบข้อมูล User
app.get('/delete/:check_ID', async (request, response) => {

    if (request.params.check_ID) {
        var sql = "DELETE FROM userpass WHERE ID = '" + request.params.check_ID + "'";
        Delete_User = await get_mysql_data(sql) // ลบข้อมูลผ่าน ID ที่ถูก request
        response.redirect('/UserPass_Info')
    }
})


/************************************* สำหรับ จัดการข้อมูลการยืมสินค้า ************************************************* */
app.all('/Departure_Oxygen_Concentration', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Departure/Departure_Oxygen_Concentration')
    }
    else if (request.method == 'POST') {

        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Departure", "Oxygen Concentration")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Return_Oxygen_Concentration', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Return/Return_Oxygen_Concentration')
    }
    else if (request.method == 'POST') {
        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Return", "Oxygen Concentration")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Departure_Suction_Machine', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Departure/Departure_Suction_Machine')
    }
    else if (request.method == 'POST') {

        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Departure", "Suction Machine")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Return_Suction_Machine', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Return/Return_Suction_Machine')
    }
    else if (request.method == 'POST') {
        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Return", "Suction Machine")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Departure_Blood_Pressure', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Departure/Departure_Blood_Pressure')
    }
    else if (request.method == 'POST') {

        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Departure", "Blood Pressure Monitor")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Return_Blood_Pressure', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Return/Return_Blood_Pressure')
    }
    else if (request.method == 'POST') {
        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Return", "Blood Pressure Monitor")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Departure_Surgical_Light', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Departure/Departure_Surgical_Light')
    }
    else if (request.method == 'POST') {

        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Departure", "Surgical Light")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Return_Surgical_Light', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Return/Return_Surgical_Light')
    }
    else if (request.method == 'POST') {
        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Return", "Surgical Light")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Departure_Surgery_Unit', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Departure/Departure_Surgery_Unit')
    }
    else if (request.method == 'POST') {

        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Departure", "Surgery Unit")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Return_Surgery_Unit', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Return/Return_Surgery_Unit')
    }
    else if (request.method == 'POST') {
        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Return", "Surgery Unit")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Departure_Infusion_Pump', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Departure/Departure_Infusion_Pump')
    }
    else if (request.method == 'POST') {

        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Departure", "Infusion Pump")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Return_Infusion_Pump', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Return/Return_Infusion_Pump')
    }
    else if (request.method == 'POST') {
        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Return", "Infusion Pump")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Departure_Patient_Monitor', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Departure/Departure_Patient_Monitor')
    }
    else if (request.method == 'POST') {

        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Departure", "Patient Monitor")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Return_Patient_Monitor', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Return/Return_Patient_Monitor')
    }
    else if (request.method == 'POST') {
        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Return", "Patient Monitor")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Departure_Vascular_Doppler', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Departure/Departure_Vascular_Doppler')
    }
    else if (request.method == 'POST') {

        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Departure", "Vascular Doppler")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Return_Vascular_Doppler', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Return/Return_Vascular_Doppler')
    }
    else if (request.method == 'POST') {
        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Return", "Vascular Doppler")
       console.log(UpdateWarehouse) 
    }
})


app.all('/Departure_Laryngoscope', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Departure/Departure_Laryngoscope')
    }
    else if (request.method == 'POST') {

        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Departure", "Laryngoscope")
       console.log(UpdateWarehouse) 
    }
})

app.all('/Return_Laryngoscope', async (request, response) => {

    if (request.method == 'GET') {
        response.render('Return/Return_Laryngoscope')
    }
    else if (request.method == 'POST') {
        // เก็บข้อมูลจากหน้าเว็บ
       var UpdateWarehouse = await Update_Warehouse_Stocklist_Googlesheet(request, response, "Return", "Laryngoscope")
       console.log(UpdateWarehouse) 
    }
})


app.all('/success', function (request, response) {

    response.render('success')
})

app.all('/error', function (request, response) {

    response.render('error')
})


app.all('/Departure-return', function (request, response) {
    response.render('departure-return')
})


app.all('/Product_line_departure', async function (request, response) {
    response.render('Product_line_departure')
})

app.all('/Product_line_return', async function (request, response) {
    response.render('Product_line_return')
})


app.listen(3000, () => console.log('Server started on port: 3000'))