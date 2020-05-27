//
//  XlsxDecoder.swift
//  AutoLocalizable
//
//  Created by Chancc on 2020/5/8.
//  Copyright © 2020 Chancc. All rights reserved.
//

import UIKit
import CoreXLSX

/*
 
 使用方法，请在电脑模拟器上跑 （README1、README2）
 
 1、将 1.xlsx 复制到桌面，这是一个对照表做了如下操作
 * 全选（cmd + a），清除表格所有内容的格式
 * 只能有一张（Sheet1）表格
 * 清除表格多余的行与列(没用到的行与列)，不清除也没关系
 2、设置输入文件路径 let sourcePath = "/Users/xxx/Desktop/1.xlsx"
 3、设置输出路径 let wirtePath = "/Users/xxx/Desktop/"
 */

/// 输入文件路径
let sourcePath = "/Users/zowell/Desktop/1.xlsx"
/// 输出路径
let wirtePath = "/Users/zowell/Desktop/"

protocol XlsxDecoder {
    func decoder()
}

extension XlsxDecoder {
    
    /// 解析 xlsx 生成 .strings 文件
    func decoder() {
        guard let file = XLSXFile(filepath: sourcePath) else {
            fatalError("XLSX file corrupted or does not exist")
        }
        guard let sharedStrings = try? file.parseSharedStrings() else {
            fatalError("SharedStrings ❌❌❌")
        }
        let path = try? file.parseWorksheetPaths()
        let worksheet = try? file.parseWorksheet(at: path?.first ?? "")
        
        guard let lastColumn = worksheet?.data?.rows.last?.cells.last?.reference.column.value, let lastRow = worksheet?.data?.rows.last?.cells.last?.reference.row else {
            fatalError("Column && Row ❌❌❌")
        }
        
        let columnInt = (lastColumn as NSString).character(at: 0)
        for column in 65...columnInt {
            let columnString = String(format: "%c", column)
            if columnString == "A" { continue }
            
            var lanString = worksheet?.cells(atColumns: [ColumnReference(String(format: "%c", column))!], rows: [1]).first?.stringValue(sharedStrings)
            print("------ \(lanString ?? "❌") ------")
            /// 文件名 lanString 不规范，规范后可以直接写出指定文件名输出
            if let a = lanString?.contains("/"), a{
                lanString = lanString?.replacingOccurrences(of: "/", with: "")
            }
            let fileURL = URL(fileURLWithPath: wirtePath + "\(lanString ?? columnString).strings")
            var data: Data? = Data()
            data?.append("//\(lanString ?? "") \r\n".data(using: .utf8) ?? Data())
            
            for row in 1...lastRow {
                let IDString = worksheet?.cells(atColumns: [ColumnReference("A")!], rows: [row]).first?.stringValue(sharedStrings)
                let valueString = worksheet?.cells(atColumns: [ColumnReference(columnString)!], rows: [row]).first?.stringValue(sharedStrings)
                guard IDString != valueString, row != 1 else {
                    //同名字符串：
                    //if row != 1 { print("\"\(IDString ?? "")\" = \"\(valueString ?? "")\";")}
                    continue
                }
                data?.append("\"\(IDString ?? "❌")\" = \"\(rules(valueString) ?? "❌")\";\r\n".data(using: .utf8) ?? Data())
            }
            try? data?.write(to: fileURL, options: .atomicWrite)
            print("🎉🎉🎉 :  \(lanString ?? "❌") \((lanString != nil) ? "😀" : "😭")")
        }
    }
    
    /// 字符串规则匹配
    func rules(_ string: String?) -> String? {
        var value = string
        // 有 " ' 加转义符号 \
        if let a = value?.contains("\""), a {
            if let a = value?.contains("\\\""), a {
                value = value?.replacingOccurrences(of: "\\\"", with: "\\\"")
            } else {
                value = value?.replacingOccurrences(of: "\"", with: "\\\"")
            }
        }
        if let a = string?.contains("\'"), a {
            if let a = value?.contains("\\\'"), a {
                value = value?.replacingOccurrences(of: "\\\'", with: "\\\'")
            } else {
                value = value?.replacingOccurrences(of: "\'", with: "\\\'")
            }
        }
        return value
    }
}
