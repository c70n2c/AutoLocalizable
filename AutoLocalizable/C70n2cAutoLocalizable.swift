//
//  C70n2cAutoLocalizable.swift
//  AutoLocalizable
//
//  Created by c70n2c on 2020/5/7.
//  Copyright © 2020 Chancc. All rights reserved.
//
import UIKit
import CoreXLSX

/// 输入文件路径
let sourcePath = "/Users/zowell/Desktop/222.xlsx"
/// 输出路径
let wirtePath = "/Users/zowell/Desktop/"

class C70n2cAutoLocalizable: NSObject {
    
    public static func xlsxDecoder() {
        guard let file = XLSXFile(filepath: sourcePath) else {
            fatalError("XLSX file corrupted or does not exist")
        }
        let path = try? file.parseWorksheetPaths()
        let worksheet = try? file.parseWorksheet(at: path?.first ?? "")
        guard let sharedStrings = try? file.parseSharedStrings() else {
            fatalError("sharedStrings ❌❌❌")
        }
        
        var newValue: [CellReference] = []
        for row in worksheet?.data?.rows ?? [] {
            for c in row.cells {
                if let _ = c.stringValue(sharedStrings) {
                    newValue += [c.reference]
                }
            }
        }
        
        guard let lastColumn = newValue.last?.column.value, let lastRow = newValue.last?.row else {
            fatalError("Column && Row ❌❌❌")
        }
        let columnInt = (lastColumn as NSString).character(at: 0)
        for column in 65...columnInt {
            let columnString = String(format: "%c", column)
            if columnString == "A" { continue }
            
            let lanString = worksheet?.cells(atColumns: [ColumnReference(String(format: "%c", column))!], rows: [1]).first?.stringValue(sharedStrings)
            print("==================\(lanString ?? "❌")==================")
            
            /// 文件名 lanString 不规范，规范后可以直接写出指定文件名输出
            let fileURL = URL(fileURLWithPath: wirtePath + "\(column).strings")
            var data: Data? = Data()
            data?.append("//\(lanString ?? "") \r\n".data(using: .utf8) ?? Data())
            
            for row in 1...lastRow {
                let IDString = worksheet?.cells(atColumns: [ColumnReference("A")!], rows: [row]).first?.stringValue(sharedStrings)
                var valueString = worksheet?.cells(atColumns: [ColumnReference(columnString)!], rows: [row]).first?.stringValue(sharedStrings)
                if IDString != valueString, row != 1 {
                    if let a = valueString?.contains("\""), a {
                        if let a = valueString?.contains("\\\""), a {
                            valueString = valueString?.replacingOccurrences(of: "\\\"", with: "\\\"")
                        } else {
                            valueString = valueString?.replacingOccurrences(of: "\"", with: "\\\"")
                        }
                    }
                    
                    if let a = valueString?.contains("\'"), a {
                        if let a = valueString?.contains("\\\'"), a {
                            valueString = valueString?.replacingOccurrences(of: "\\\'", with: "\\\'")
                        } else {
                            valueString = valueString?.replacingOccurrences(of: "\'", with: "\\\'")
                        }
                    }
                    let finalString = "\"\(IDString ?? "❌")\" = \"\(valueString ?? "❌")\";\r\n"
                    data?.append(finalString.data(using: .utf8) ?? Data())
                    try? data?.write(to: fileURL, options: .atomicWrite)
                }
            }
            print("😀😀😀😀😀😀😀😀😀😀😀 结束")
        }
    }
    
    
    
    //csv , 太多 有些字符串无法判断，因此写的有问题
    func csv() {
        //let fileContent = try? String(contentsOfFile: Bundle.main.path(forResource: fileName, ofType: "") ?? "")
        let fileContent = try? String(contentsOfFile: sourcePath)
        if var values = fileContent?.components(separatedBy: ",,,,,,,,,,,\r\n"), var languages = values.first?.components(separatedBy: ",") {
            values.removeAll(where: {$0 == ""})
            values.removeFirst()
            var newValues: [[String]] = []
            for value in values {
                var keyValues = value.components(separatedBy: ",")
                keyValues.removeAll(where: {$0 == ""})
                newValues += [keyValues]
            }
            languages.removeFirst()
            languages.removeAll(where: {$0 == ""})
            for (i, language) in languages.enumerated() {
                for (j, values) in newValues.enumerated() where values.first ?? "" != newValues[j][i] {
                    
                    if newValues[j][i].contains("\"") {
                        newValues[j][i].removeAll(where: {$0 == "\""})
                    }
                    if newValues[j][i].contains("\n") {
                        newValues[j][i].removeAll(where: {$0 == "\n"})
                    }
                    print("\"\(values.first ?? "❌")\" = \"\(newValues[j][i])\";")
                }
                print("====================== \(language) ======================")
            }
        }
    }
    
    
    
    
}

