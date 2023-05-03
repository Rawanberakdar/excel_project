import 'package:flutter/material.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'dart:io';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';
import 'package:universal_html/html.dart' show AnchorElement;
import 'package:flutter/foundation.dart' show kIsWeb;
import 'dart:convert';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Project',
      theme: ThemeData(
        // This is the theme of your application.
        //
        // Try running your application with "flutter run". You'll see the
        // application has a blue toolbar. Then, without quitting the app, try
        // changing the primarySwatch below to Colors.green and then invoke
        // "hot reload" (press "r" in the console where you ran "flutter run",
        // or simply save your changes to "hot reload" in a Flutter IDE).
        // Notice that Color.fromARGB(255, 197, 211, 222) didn't reset back to zero; the application
        // is not restarted.
        primarySwatch: Colors.blue,
      ),
      home: const MyHomePage(title: 'Flutter excel project'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({Key? key, required this.title}) : super(key: key);

  final String title;

  @override
  // ignore: library_private_types_in_public_api
  _MyHomePageState createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: Color.fromARGB(255, 94, 118, 131),
      body: Center(
        child: ElevatedButton(
            onPressed: createExcel,
            child: const Text('CLICK HERE TO VIEW EXCEL FILE ')),
      ),
    );
  }

  Future<void> createExcel() async {
    final Workbook workbook = Workbook();
    final Worksheet sheet = workbook.worksheets[0];

    final List<int> imageBytes = File('images/image.png').readAsBytesSync();
    //sheet.pictures.addStream(3, 10, imageBytes);
    final Picture picture = sheet.pictures.addStream(1, 15, imageBytes);
    picture.height = 150;
    picture.width = 150;

    sheet.name = 'Student';
    final Worksheet sheet2 = workbook.worksheets.addWithName('CLOs');
    final Worksheet sheet3 = workbook.worksheets.addWithName('Quizzes');
    final Worksheet sheet4 = workbook.worksheets.addWithName('Assignments');

    Style globalStyle = workbook.styles.add('style');
    globalStyle.backColor = '#b9c3c4';
    globalStyle.fontName = 'Times New Roman';
    globalStyle.fontSize = 20;
    globalStyle.fontColor = '#000000';
    globalStyle.italic = true;
    globalStyle.bold = true;
    globalStyle.borders.all.color = '#000000';

    final Style style1 = workbook.styles.add('Style1');
    style1.fontName = 'Times New Roman';
    style1.backColor = '#276496';
    style1.bold = true;
    style1.fontSize = 15;
    style1.fontColor = '#FFFFFF';
    style1.vAlign = VAlignType.center;
    style1.italic = true;
    style1.borders.bottom.lineStyle = LineStyle.thin;
    style1.borders.bottom.color = '#A6A6A6';
    style1.borders.right.lineStyle = LineStyle.thick;
    style1.borders.right.color = '#A6A6A6';

    final Style style2 = workbook.styles.add('Style2');
    style2.fontName = 'Times New Roman';
    style2.backColor = '#d9c448';
    style2.bold = true;
    style2.fontSize = 15;
    style2.fontColor = '#12110e';
    style2.vAlign = VAlignType.center;

    style2.borders.bottom.lineStyle = LineStyle.thin;
    style2.borders.bottom.color = '#A6A6A6';
    style2.borders.right.lineStyle = LineStyle.thick;
    style2.borders.right.color = '#A6A6A6';

    //set the side in same color
    sheet.getRangeByName('A1:A35').cellStyle = globalStyle;
    sheet.getRangeByName('A1:S1').cellStyle = globalStyle;
    sheet.getRangeByName('S1:S25').cellStyle = globalStyle;

    sheet.getRangeByName('B2:l2').cellStyle.backColor = '#276496';
    sheet.getRangeByName('B2:l2').merge();
    sheet
        .getRangeByName('B2:l2')
        .setText('College of Computer Science & Information System');
    sheet.getRangeByName('B2:l2').cellStyle.fontColor = '#faf8ed';
    sheet.getRangeByName('B2:l2').cellStyle.fontSize = 25;

    sheet.getRangeByName('M2:R2').cellStyle.backColor = '#bfa928';
    sheet.getRangeByName('M3:R3').cellStyle.backColor = '#bfa928';
    sheet.getRangeByName('B3:l3').cellStyle.backColor = '#276496';

    sheet.getRangeByName('B4:R4').cellStyle = globalStyle;
    sheet.getRangeByName('B4:G4').merge();

    sheet.getRangeByName('B4').setText('Student Information');

    sheet.getRangeByName('B5:N5').merge();
    sheet.getRangeByName('B5').setText(
        ' Department : Information System     Class:BSc     Semester: First Semester 2022/2023');
    sheet.getRangeByName('B6:N6').merge();
    sheet.getRangeByName('B6').setText(
        ' Section : 55   Course Code:415CIS-3  Course Title: Mobile Application Developement');
    sheet.getRangeByName('B7:N7').merge();
    sheet.getRangeByName('B7').setText(
        'credit Hourse : 3        No.of Student : 24       Instructor : DR.Mohamed Khairi ');
    sheet.getRangeByName('B5:R5').cellStyle.backColor = '#78b8c6';
    sheet.getRangeByName('B6:R6').cellStyle.backColor = '#78b8c6';
    sheet.getRangeByName('B7:R7').cellStyle.backColor = '#78b8c6';

    sheet.getRangeByName('B8:R8').cellStyle = globalStyle;

    sheet.getRangeByName('B9:F9').cellStyle = style1;
    //sheet.getRangeByName('B9:F9').cellStyle.backColor = '#276496';

    sheet.getRangeByName('B9').setText('S.NO');

    sheet.getRangeByName('C9:D9').merge();

    sheet.getRangeByName('C9:D9').setText('StudentID');
    sheet.getRangeByName('C9:F9').cellStyle.fontColor = '#faf8ed';
    sheet.getRangeByName('E9:F9').merge();
    sheet.getRangeByName('E9:F9').setText('StudentName');

    sheet.getRangeByName('B9:B33').cellStyle = style1;

    sheet.getRangeByName('B10').setNumber(1);
    sheet.getRangeByName('B11').setNumber(2);
    sheet.getRangeByName('B12').setNumber(3);
    sheet.getRangeByName('B13').setNumber(4);
    sheet.getRangeByName('B14').setNumber(5);
    sheet.getRangeByName('B15').setNumber(6);
    sheet.getRangeByName('B16').setNumber(7);
    sheet.getRangeByName('B17').setNumber(8);
    sheet.getRangeByName('B18').setNumber(9);
    sheet.getRangeByName('B19').setNumber(10);
    sheet.getRangeByName('B20').setNumber(11);
    sheet.getRangeByName('B21').setNumber(12);
    sheet.getRangeByName('B22').setNumber(13);
    sheet.getRangeByName('B23').setNumber(14);
    sheet.getRangeByName('B24').setNumber(15);
    sheet.getRangeByName('B25').setNumber(16);
    sheet.getRangeByName('B26').setNumber(17);
    sheet.getRangeByName('B27').setNumber(18);
    sheet.getRangeByName('B28').setNumber(19);
    sheet.getRangeByName('B29').setNumber(20);
    sheet.getRangeByName('B30').setNumber(21);
    sheet.getRangeByName('B31').setNumber(22);
    sheet.getRangeByName('B32').setNumber(23);
    sheet.getRangeByName('B33').setNumber(24);

    sheet.getRangeByName('C10:D10').merge();
    sheet.getRangeByName('C10').setNumber(2334456);
    sheet.getRangeByName('C11:D11').merge();
    sheet.getRangeByName('C11').setNumber(2334456);
    sheet.getRangeByName('C12:D12').merge();
    sheet.getRangeByName('C12').setNumber(2334456);
    sheet.getRangeByName('C13:D13').merge();
    sheet.getRangeByName('C13').setNumber(2334456);
    sheet.getRangeByName('C14:D14').merge();
    sheet.getRangeByName('C14').setNumber(2767456);
    sheet.getRangeByName('C15:D15').merge();
    sheet.getRangeByName('C15').setNumber(2330458);
    sheet.getRangeByName('C16:D16').merge();
    sheet.getRangeByName('C16').setNumber(2357456);
    sheet.getRangeByName('C17:D17').merge();
    sheet.getRangeByName('C17').setNumber(2334456);
    sheet.getRangeByName('C18:D18').merge();
    sheet.getRangeByName('C18').setNumber(2334459);
    sheet.getRangeByName('C19:D19').merge();
    sheet.getRangeByName('C19').setNumber(2334455);
    sheet.getRangeByName('C20:D20').merge();
    sheet.getRangeByName('C20').setNumber(2334455);
    sheet.getRangeByName('C21:D21').merge();
    sheet.getRangeByName('C21').setNumber(2334455);
    sheet.getRangeByName('C22:D22').merge();
    sheet.getRangeByName('C22').setNumber(2334455);
    sheet.getRangeByName('C23:D23').merge();
    sheet.getRangeByName('C23').setNumber(2334455);
    sheet.getRangeByName('C24:D24').merge();
    sheet.getRangeByName('C24').setNumber(2334455);
    sheet.getRangeByName('C25:D25').merge();
    sheet.getRangeByName('C25').setNumber(2334455);
    sheet.getRangeByName('C26:D26').merge();
    sheet.getRangeByName('C26').setNumber(2334455);
    sheet.getRangeByName('C27:D27').merge();
    sheet.getRangeByName('C27').setNumber(2334455);
    sheet.getRangeByName('C28:D28').merge();
    sheet.getRangeByName('C28').setNumber(2334455);
    sheet.getRangeByName('C29:D29').merge();
    sheet.getRangeByName('C29').setNumber(2334455);
    sheet.getRangeByName('C30:D30').merge();
    sheet.getRangeByName('C30').setNumber(2334455);
    sheet.getRangeByName('C31:D31').merge();
    sheet.getRangeByName('C31').setNumber(2334455);
    sheet.getRangeByName('C32:D32').merge();
    sheet.getRangeByName('C32').setNumber(2334455);
    sheet.getRangeByName('C33:D33').merge();
    sheet.getRangeByName('C33').setNumber(2334455);

    sheet.getRangeByName('E10:F10').merge();
    sheet.getRangeByName('E10').setText('علي مسفربن حسين');
    sheet.getRangeByName('E11:F11').merge();
    sheet.getRangeByName('E11').setText('محمد فهد الغامدي');
    sheet.getRangeByName('E12:F12').merge();
    sheet.getRangeByName('E12').setText('ابراهيم فارس بن محمد');
    sheet.getRangeByName('E13:F13').merge();
    sheet.getRangeByName('E13').setText('حسين مسفر القحطاني');
    sheet.getRangeByName('E14:F14').merge();
    sheet.getRangeByName('E14').setText('  ابراهيم بن علي رحابي ');
    sheet.getRangeByName('E15:F15').merge();
    sheet.getRangeByName('E15').setText('الوليد خالد ال راشج');
    sheet.getRangeByName('E16:F16').merge();
    sheet.getRangeByName('E16').setText('محمد ابراهيم ماجد');
    sheet.getRangeByName('E17:F17').merge();
    sheet.getRangeByName('E17').setText('مانع علي مسفر');
    sheet.getRangeByName('E18:F18').merge();
    sheet.getRangeByName('E18').setText('محمد اسحاق مصطفى ');
    sheet.getRangeByName('E19:F19').merge();
    sheet.getRangeByName('E19').setText('محمد سالم ال سالم');
    sheet.getRangeByName('E20:F20').merge();
    sheet.getRangeByName('E20').setText(' يوسف بن غنيم');
    sheet.getRangeByName('E21:F21').merge();
    sheet.getRangeByName('E21').setText('   عبد العزيز بن مهدي');
    sheet.getRangeByName('E22:F22').merge();
    sheet.getRangeByName('E22').setText('  احمد بن طالب ');
    sheet.getRangeByName('E23:F23').merge();
    sheet.getRangeByName('E23').setText('حاتم بن ابراهيم ');
    sheet.getRangeByName('E24:F24').merge();
    sheet.getRangeByName('E24').setText('محمد سالم ال سالم');
    sheet.getRangeByName('E25:F25').merge();
    sheet.getRangeByName('E25').setText('محمد عامر ال متعب');
    sheet.getRangeByName('E26:F26').merge();
    sheet.getRangeByName('E26').setText('محمد سالم ال سالم');
    sheet.getRangeByName('E27:F27').merge();
    sheet.getRangeByName('E27').setText('عيسى بن ال سالم');
    sheet.getRangeByName('E28:F28').merge();
    sheet.getRangeByName('E28').setText('خالد صالح عبد العزيز ');
    sheet.getRangeByName('E29:F29').merge();
    sheet.getRangeByName('E29').setText('مانع علي بن مسفر');
    sheet.getRangeByName('E30:F30').merge();
    sheet.getRangeByName('E30').setText('فراس مسعود ال مصلح');
    sheet.getRangeByName('E31:F31').merge();
    sheet.getRangeByName('E31').setText('رياض سعد ال سعود');
    sheet.getRangeByName('E32:F32').merge();
    sheet.getRangeByName('E32').setText('محمد سالم مفرح ');
    sheet.getRangeByName('E33:F33').merge();
    sheet.getRangeByName('E33').setText('حسن بن ابراهيم');

    // final ExcelTable table =
    //     sheet.tableCollection.create('Table1', sheet.getRangeByName('B9:E19'));
    // table.builtInTableStyle = ExcelTableBuiltInStyle.tableStyleDark10;
    // table.showFirstColumn = true;
    // sheet.getRangeByName('G9:G33').merge();
    sheet.getRangeByName('G9:G19').cellStyle = globalStyle;
    sheet.getRangeByName('H9:R9').cellStyle = globalStyle;
    sheet.getRangeByName('G10:R10').cellStyle = globalStyle;
    //name of student who make the project
    sheet.getRangeByName('I9:N9').merge();
    sheet.getRangeByName('I9:N9').cellStyle = style1;
    sheet
        .getRangeByName('I9')
        .setText('    تطوير الاجهزة المحمولة بإشراف الدكتور:محمد خيري');
    sheet.getRangeByName('I10:N10').merge();
    sheet.getRangeByName('I10:N10').cellStyle = style1;

    sheet.getRangeByName('I10').setText('          إعداد الطلاب        ');
    sheet.getRangeByName('I11:N11').merge();
    sheet.getRangeByName('I11:N11').cellStyle = style2;
    sheet.getRangeByName('I11').setText('  سياف حسن ناصر عسيري  439100481');

    sheet.getRangeByName('I12:N12').merge();
    sheet.getRangeByName('I12:N12').cellStyle = style2;
    sheet
        .getRangeByName('I12')
        .setText('  ابراهيم عبد الله علي الوالي 441106681 ');
    sheet.getRangeByName('I13:N13').merge();
    sheet.getRangeByName('I13:N13').cellStyle = style2;
    sheet.getRangeByName('I13').setText('مهدي ابراهيم مهدي ال جعفر 441207675 ');
    sheet.getRangeByName('G9:H33').merge();
    sheet.getRangeByName('G9:H33').cellStyle = globalStyle;
    sheet.getRangeByName('I14:S33').merge();
    sheet.getRangeByName('I14:S33').cellStyle = globalStyle;
    sheet.getRangeByName('O11:R13').merge();
    sheet.getRangeByName('O11:R13').cellStyle = globalStyle;
///////////////////////////////////SHEET2/////////////////////////////
//sheet2___IF_THAT_INCLUDE__//
    sheet2.getRangeByName('A1:A35').cellStyle = globalStyle;
    sheet2.getRangeByName('B1:T1').cellStyle = globalStyle;
    final List<int> imageBytes2 = File('images/image.png').readAsBytesSync();
    //sheet.pictures.addStream(3, 10, imageBytes);
    final Picture picture2 = sheet2.pictures.addStream(2, 15, imageBytes);
    picture2.height = 100;
    picture2.width = 100;

    sheet2.getRangeByName('B2:l2').cellStyle.backColor = '#276496';
    sheet2.getRangeByName('B2:l2').merge();
    sheet2
        .getRangeByName('B2:l2')
        .setText('College of Computer Science & Information System');
    sheet2.getRangeByName('B2:l2').cellStyle.fontColor = '#faf8ed';
    sheet2.getRangeByName('B2:l2').cellStyle.fontSize = 25;

    sheet2.getRangeByName('M2:R2').cellStyle.backColor = '#bfa928';
    sheet2.getRangeByName('M3:R3').cellStyle.backColor = '#bfa928';
    sheet2.getRangeByName('B3:l3').cellStyle.backColor = '#276496';
    sheet2.getRangeByName('B4:R4').cellStyle = globalStyle;
    sheet2.getRangeByName('B4:G4').merge();
    sheet2.getRangeByName('B4').setText('Course Learning Outcomes CLOs');

    sheet2.getRangeByName('B5:N5').merge();
    sheet2.getRangeByName('B5').setText(
        ' Department : Information System     Class:BSc     Semester: First Semester 2022/2023');
    sheet2.getRangeByName('B6:N6').merge();
    sheet2.getRangeByName('B6').setText(
        ' Section : 55   Course Code:415CIS-3  Course Title: Mobile Application Developement');
    sheet2.getRangeByName('B7:N7').merge();
    sheet2.getRangeByName('B7').setText(
        'credit Hourse : 3        No.of Student : 24       Instructor : DR.Mohamed Khairi ');
    sheet2.getRangeByName('B5:R5').cellStyle.backColor = '#78b8c6';
    sheet2.getRangeByName('B6:R6').cellStyle.backColor = '#78b8c6';
    sheet2.getRangeByName('B7:R7').cellStyle.backColor = '#78b8c6';

    sheet2.getRangeByName('B8:R8').cellStyle = globalStyle;

////////////////////////////////////////////////////
    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    if (kIsWeb) {
      AnchorElement(
          href:
              'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}')
        ..setAttribute('download', 'Output.xlsx')
        ..click();
    } else {
      final String path = (await getApplicationSupportDirectory()).path;
      final String fileName =
          Platform.isWindows ? '$path\\Output.xlsx' : '$path/Output.xlsx';
      final File file = File(fileName);
      await file.writeAsBytes(bytes, flush: true);
      OpenFile.open(fileName);
    }
  }
}
