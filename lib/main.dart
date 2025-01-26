import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:desktop_drop/desktop_drop.dart';
import 'package:cross_file/cross_file.dart';
import 'package:flutter/services.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const MyHomePage(),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key});

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  // late OverlayEntry overlayEntry;
  final _directoryController = TextEditingController();
  final _rowController = TextEditingController();
  final _columnController = TextEditingController();
  final _sheetController = TextEditingController();
  String tipStr = '';
  List<String> sheetListData = [];
  List<String> resultList = [];
  List<String> resultTempList = [];

  @override
  void initState() {
    super.initState();
    // WidgetsBinding.instance.addPostFrameCallback((timeStamp) {
    //   overlayEntry = OverlayEntry(builder: (context) {
    //     return Positioned(child: Container(child:Text('hello')));
    //   });
    // });
    _directoryController.text = '';
    // _sheetController.text = '财务指标';
    _rowController.text = '1';
    _columnController.text = '2';
    readExcelSheet(_directoryController.text);
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: const Text('READ EXCEL'),
      ),
      body: Container(
        width: MediaQuery.of(context).size.width,
        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 16),
        child: Column(
          mainAxisAlignment: MainAxisAlignment.start,
          children: <Widget>[
            buildRow(
                controller: _directoryController,
                hindText: '选择需要读取的Excel',
                onButtonClick: () async {
                  String? dir = await pickerExcelFilePath();
                  var path = processPickedPathOrDir(dir);
                  readExcelSheet(path);
                },
                onChanged: (value) {
                  setState(() {
                    _directoryController.text = value;
                  });
                  // _refreshCursor(_directoryController);
                },
                onDragDone: (files) {
                  var path = files.first.path;
                  readExcelSheet(path);
                }),
            buildSheetField(
                controller: _sheetController,
                hindText: '请输入sheet名称',
                onChanged: (value) {
                  setState(() {
                    _sheetController.text = value.trim();
                  });
                }),
            buildRowColumn(
                controller: _rowController,
                hindText: '请输入从第几行读起,不输入则从0开始',
                onChanged: (value) {
                  setState(() {
                    _rowController.text = value.trim();
                  });
                }),
            buildRowColumn(
                controller: _columnController,
                hindText: '请输入从第几列读起,不输入则从0开始',
                onChanged: (value) {
                  setState(() {
                    _columnController.text = value.trim();
                  });
                }),
            tipStr.isNotEmpty
                ? Container(
                    margin: const EdgeInsets.only(top: 5),
                    child: Text(
                      tipStr,
                      style: const TextStyle(color: Colors.red),
                    ),
                  )
                : const SizedBox(),
            Container(
              margin: const EdgeInsets.only(top: 20),
              child: ElevatedButton(
                  onPressed: () async {
                    var path = _directoryController.text.trim();
                    if (path.isEmpty) {
                      setState(() {
                        tipStr = '请先输入正确路径！';
                      });
                      return;
                    }
                    if (!path.endsWith('.xls') && !path.endsWith('.xlsx')) {
                      setState(() {
                        tipStr = '请选择EXCEL文件！';
                      });
                      return;
                    }
                    var file = File(path);
                    if (await file.exists()) {
                      setState(() {
                        tipStr = '';
                      });
                      readExcel(file);
                    } else {
                      setState(() {
                        tipStr = '当前路径不存在！';
                      });
                    }
                  },
                  child: const Text('读取数据并转换')),
            ),
            const SizedBox(
              height: 20,
            ),
            resultList.isNotEmpty
                ? Row(
                    mainAxisAlignment: MainAxisAlignment.end,
                    children: [
                      ElevatedButton(
                        onPressed: () {
                          Clipboard.setData(
                                  ClipboardData(text: resultList.toString()))
                              .then((value) {
                            showToast(context, '复制成功');
                          });
                        },
                        style: ButtonStyle(
                            backgroundColor: MaterialStateProperty.all(
                                Colors.grey.shade300)),
                        child: const Text('复制结果'),
                      )
                    ],
                  )
                : const SizedBox(),
            resultList.isNotEmpty
                ? Expanded(
                    child: Container(
                    margin: const EdgeInsets.only(top: 5),
                    width: MediaQuery.of(context).size.width,
                    decoration: BoxDecoration(
                        color: Colors.grey.shade300,
                        borderRadius:
                            const BorderRadius.all(Radius.circular(10))),
                    child: SingleChildScrollView(
                      child: Container(
                        padding: const EdgeInsets.all(16),
                        child: Text('$resultList'),
                      ),
                    ),
                  ))
                : const SizedBox()
          ],
        ),
      ), // This trailing comma makes auto-formatting nicer for build methods.
    );
  }

  void readExcel(File file) {
    var rowIndex = 0;
    if (_rowController.text.isNotEmpty) {
      rowIndex = int.parse(_rowController.text);
    }
    var columnIndex = 0;
    if (_columnController.text.isNotEmpty) {
      columnIndex = int.parse(_columnController.text);
    }
    var bytes = file.readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    var sheetName = _sheetController.text.trim();
    if (sheetName.isNotEmpty) {
      if (excel.tables.containsKey(sheetName)) {
        var _sheet = excel.tables[sheetName]!;
        readRowAndColumn(sheetName, _sheet, rowIndex, columnIndex);
      } else {
        setState(() {
          tipStr = '表格中不存在所输入sheet！';
        });
      }
    } else {
      for (var table in excel.tables.keys) {
        print(table);
        readRowAndColumn(table, excel.tables[table]!, rowIndex, columnIndex);
      }
    }
    // for (var table in excel.tables.keys) {
    //   print(table);
    //   //sheet Name
    //   print(excel.tables[table]?.maxCols);
    //   print(excel.tables[table]?.maxRows);
    //   for (var row in excel.tables[table]!.rows) {
    //     for (var x in row) {
    //       if (x != null) {
    //         var xx = x.value;
    //         print("xx==$xx,${x.rowIndex},${x.colIndex}");
    //       }
    //     }
    //   }
    // }
  }

  void readRowAndColumn(
      String sheetName, Sheet sheet, int rowIndex, int columnIndex) {
    resultList.clear();
    if (sheet.rows.length > rowIndex) {
      for (int i = rowIndex; i < sheet.rows.length; i++) {
        var columns = sheet.rows[i];
        if (columns.length > 3) {
          //读取财报
          if (columns[3] != null && columns[2] != null) {
            var keyTemp = columns[3]!.value.toString().trim();
            List<String> words = keyTemp.split('_');
            // debugPrint('words=$words,$keyTemp');
            words = words.map((word) {
              if (words.indexOf(word) == 0) {
                return word;
              } else {
                return word[0].toUpperCase() + word.substring(1);
              }
            }).toList();
            var map = {
              words.join(''): columns[2]!.value.toString().trim(),
              'type': 'content',
              'level': '2'
            };
            resultList.add(jsonEncode(map));
          } else {
            var map = {
              'titleTemp': '${sheetName}Sheet${i + 1}行需修改title',
              'type': 'title',
              'level': '1'
            };
            resultList.add(jsonEncode(map));
          }
        }
        if (columns.length > columnIndex) {
          for (int j = columnIndex; j < columns.length; j++) {
            if (columns[j] != null) {
              var v = columns[j]!.value;
              // print('v=$rowIndex=====$v');
              var map = {'${i + 1}行${getColumnLetter(j + 1)}${j + 1}列': '$v'};
              resultTempList.add(jsonEncode(map));
            } else {}
          }
        } else {
          for (int j = 0; j < columns.length; j++) {
            if (columns[j] != null) {
              var v = columns[j]!.value;
              var map = {'${i + 1}行${getColumnLetter(j + 1)}${j + 1}列': '$v'};
              resultTempList.add(jsonEncode(map));
            }
          }
          setState(() {
            tipStr = '输入列数有误,已从0开始读数据';
          });
        }
      }
      print(resultTempList);
      print(resultList);
      setState(() {});
    } else {
      setState(() {
        tipStr = '输入行数有误,下标越界';
      });
    }
  }

  readExcelSheet(String path) async {
    sheetListData.clear();
    if (!path.endsWith('.xls') && !path.endsWith('.xlsx')) {
      setState(() {
        tipStr = '请选择EXCEL文件！';
      });
      return;
    }
    var file = File(path);
    if (await file.exists()) {
      var bytes = file.readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      for (var sheetName in excel.sheets.keys) {
        sheetListData.add(sheetName);
      }
      if (sheetListData.isNotEmpty) {
        _sheetController.text = sheetListData.first;
      }
      setState(() {
        tipStr = '';
        _directoryController.text = path;
      });
    } else {
      setState(() {
        tipStr = '当前路径不存在！';
        _directoryController.text = path;
      });
    }
  }

  Future<String?> pickerExcelFilePath() async {
    String? path = '';
    FilePickerResult? result = await FilePicker.platform.pickFiles();

    if (result != null) {
      List<PlatformFile> files = result.files;
      if (files.isNotEmpty) {
        path = files.single.path;
        debugPrint('path====$path');
      }
    } else {
      // User canceled the picker
    }
    // String? selectedDirectory = await FilePicker.platform.getDirectoryPath();
    return path;
  }

  String processPickedPathOrDir(String? pathOrDir) {
    print("pathOrDir=$pathOrDir");
    var temp = '';
    if (pathOrDir != null) {
      int pos = pathOrDir.indexOf('/Users');
      if (pos >= 0) {
        temp = pathOrDir.substring(pos);
      }
    }
    return temp;
  }

  Container buildRowColumn({
    TextEditingController? controller,
    FocusNode? focusNode,
    String hindText = '',
    Function(String value)? onChanged,
  }) {
    return Container(
      margin: const EdgeInsets.only(top: 20),
      child: Row(
        children: [
          Expanded(
              child: TextField(
            focusNode: focusNode,
            controller: controller,
            keyboardType: TextInputType.number,
            maxLines: null,
            decoration: InputDecoration(
              contentPadding:
                  const EdgeInsets.symmetric(horizontal: 10, vertical: 0.0),
              hintStyle: const TextStyle(color: Colors.grey),
              filled: true,
              enabledBorder: const OutlineInputBorder(
                  borderSide: BorderSide(color: Color(0x00FF0000)),
                  borderRadius: BorderRadius.all(Radius.circular(5))),
              hintText: hindText,
              focusedBorder: const OutlineInputBorder(
                  borderSide: BorderSide(color: Color(0x00000000)),
                  borderRadius: BorderRadius.all(Radius.circular(5))),
            ),
            onChanged: (value) {
              onChanged?.call(value);
            },
          )),
          Container(
            margin: const EdgeInsets.only(left: 16),
            child: const Text('必须为数字'),
          )
        ],
      ),
    );
  }

  Container buildSheetField({
    TextEditingController? controller,
    FocusNode? focusNode,
    String hindText = '',
    Function(String value)? onChanged,
  }) {
    return Container(
      margin: const EdgeInsets.only(top: 20),
      child: Row(
        children: [
          Expanded(
              child: TextField(
            focusNode: focusNode,
            controller: controller,
            maxLines: null,
            enabled: false,
            decoration: InputDecoration(
              contentPadding:
                  const EdgeInsets.symmetric(horizontal: 10, vertical: 0.0),
              hintStyle: const TextStyle(color: Colors.grey),
              filled: true,
              enabledBorder: const OutlineInputBorder(
                  borderSide: BorderSide(color: Color(0x00FF0000)),
                  borderRadius: BorderRadius.all(Radius.circular(5))),
              disabledBorder: const OutlineInputBorder(
                  borderSide: BorderSide(color: Color(0x00FF0000)),
                  borderRadius: BorderRadius.all(Radius.circular(5))),
              hintText: hindText,
              focusedBorder: const OutlineInputBorder(
                  borderSide: BorderSide(color: Color(0x00000000)),
                  borderRadius: BorderRadius.all(Radius.circular(5))),
            ),
            onChanged: (value) {
              onChanged?.call(value);
            },
          )),
          Container(
            margin: const EdgeInsets.only(left: 16),
            child: DropdownButtonHideUnderline(
              child: DropdownButton(
                iconEnabledColor: Colors.green,
                iconSize: 50,
                style: const TextStyle(color: Colors.green),
                value: _sheetController.text.trim(),
                onChanged: (value) {
                  if (value != null) {
                    setState(() {
                      _sheetController.text = value;
                    });
                  }
                },
                items: sheetListData.map((String item) {
                  return DropdownMenuItem(
                    value: item,
                    child: Text(item),
                  );
                }).toList(),
                // selectedItemBuilder: (context) {
                //   return sheetListData.map((String item) {
                //     return Container(
                //       color: Colors.blue,
                //       alignment: Alignment.center,
                //       child: Text(
                //         item,
                //         style: TextStyle(
                //           color: Colors.red,
                //           fontSize: 16,
                //         ),
                //       ),
                //     );
                //   }).toList();
                // },
              ),
            ),
          )

          // Container(
          //   margin: const EdgeInsets.only(left: 16),
          //   child: const Text('Excel中的sheet名称'),
          // )
        ],
      ),
    );
  }

  Row buildRow(
      {TextEditingController? controller,
      FocusNode? focusNode,
      String hindText = '',
      String buttonText = '浏览',
      Function? onButtonClick,
      Function(String value)? onChanged,
      Function(List<XFile> files)? onDragDone}) {
    return Row(
      mainAxisAlignment: MainAxisAlignment.spaceBetween,
      children: [
        Expanded(
          child: DropTarget(
            onDragDone: (detail) {
              List<XFile> files = detail.files;
              if (files.isNotEmpty) {
                onDragDone?.call(files);
              }
            },
            child: TextField(
              focusNode: focusNode,
              controller: controller,
              maxLines: null,
              decoration: InputDecoration(
                contentPadding:
                    const EdgeInsets.symmetric(horizontal: 10, vertical: 0.0),
                hintStyle: const TextStyle(color: Colors.grey),
                filled: true,
                enabledBorder: const OutlineInputBorder(
                    borderSide: BorderSide(color: Color(0x00FF0000)),
                    borderRadius: BorderRadius.all(Radius.circular(5))),
                hintText: hindText,
                focusedBorder: const OutlineInputBorder(
                    borderSide: BorderSide(color: Color(0x00000000)),
                    borderRadius: BorderRadius.all(Radius.circular(5))),
              ),
              onChanged: (value) {
                onChanged?.call(value);
              },
            ),
          ),
        ),
        Container(
          margin: const EdgeInsets.only(left: 16),
          child: ElevatedButton(
              onPressed: () {
                onButtonClick?.call();
              },
              child: Text(buttonText)),
        )
      ],
    );
  }

  String getColumnLetter(int columnNumber) {
    String letter = "";
    while (columnNumber > 0) {
      int remainder = (columnNumber - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      columnNumber = (columnNumber - 1) ~/ 26;
    }
    return letter;
  }

  // 显示Toast
  void showToast(BuildContext context, String message) {
    OverlayState overlayState = Overlay.of(context);
    OverlayEntry overlayEntry = OverlayEntry(
      builder: (BuildContext context) => Positioned(
        left: (MediaQuery.of(context).size.width - 300) / 2,
        top: MediaQuery.of(context).size.height * 0.7,
        width: 300,
        child: Material(
          color: Colors.transparent,
          child: Padding(
            padding: const EdgeInsets.symmetric(horizontal: 30),
            child: Container(
              alignment: Alignment.center,
              decoration: BoxDecoration(
                color: Colors.black.withOpacity(0.5),
                borderRadius: BorderRadius.circular(8),
              ),
              padding: const EdgeInsets.symmetric(vertical: 10, horizontal: 16),
              child: Text(
                message,
                style: const TextStyle(fontSize: 16, color: Colors.white),
              ),
            ),
          ),
        ),
      ),
    );
    overlayState.insert(overlayEntry);
    Future.delayed(const Duration(seconds: 2))
        .then((value) => overlayEntry.remove()); // 2秒后自动移除Toast
  }
}
