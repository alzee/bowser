import { useState } from "react";
import reactLogo from "./assets/react.svg";
import { invoke } from "@tauri-apps/api/tauri";
import "./App.css";
import { read, utils, write } from 'xlsx';
import { useEffect } from 'react';
import { open } from '@tauri-apps/api/dialog';
import { readDir, createDir, readBinaryFile, writeBinaryFile, BaseDirectory } from '@tauri-apps/api/fs';
import { getName, getVersion } from '@tauri-apps/api/app';
import { listen } from '@tauri-apps/api/event';

const ver = await getVersion()
const appName = await getName()
const outputDir = 'keyi'

function App() {
  const [name, setName] = useState("")
  const [msg, setMsg] = useState("")
  const [dir, setDir] = useState("")
  const [file, setFile] = useState("")

  listen<string>('tauri://file-drop', (event) => {
    setDir(event.payload[0])
  })

  async function getDir() {
    const dir = await open({directory: true})
    if (!Array.isArray(dir) && dir !== null) {
      setDir(dir)
    }
  }

  async function getFile() {
    const file = await open({directory: false})
    // const contents = await readBinaryFile(file)
    // console.log(contents)
    if (file !== null) {
      console.log(file)
      setFile(file)
    }
  }

  function isXlsx() {
    return false
  }

  function processData(file) {
    if (isXlsx(file)) {
      //
    } else {
      setMsg(`[${file}] 不是有效的xls/xlsx文件`)
    }
  }

  async function main() {
    if (file === null || file === '') {
    } else {
    }

    if (dir === null || dir === '') {
      setMsg('请选择目录')
    } else {
      try {
        const entries = await readDir(dir, { recursive: true })
        setMsg('处理中...')

        const newDir = dir + '/' + outputDir
        await createDir(newDir, { recursive: true })

        processData(file)

        const sheetName = "Sheet1"

        let ws_data0 = [
          [ "案卷级档案", "姓名", "性别", "身份证号", "政治面貌", "密集架号", "总件数", '总页数' ],
        ]


        for (const entry of entries) {
          // if entry is dir


          console.log(entry)
          let xlsxFile = entry

          if (entry.children === undefined) {
            // if entry is file
            try {
              // TODO check if is xlsx to avoid read big file
              const contents = await readBinaryFile(dir + '/' + xlsxFile.name)

              const workbook = read(contents)
              const sheeName = workbook.SheetNames[0]
              const sheet = workbook.Sheets[sheeName]
              const name = sheet.A2.v.replace('姓名：', '')
              
              const ref = sheet['!ref']
              let lastRow
              if (ref !== undefined) {
                lastRow = Number(ref.replace(/[A-Z]/g,'').split(':')[1])
              } else {
                lastRow = 0
              }
              
              let docs = []
              let startRow = 6
              let sum = 0
              let count = 0

              for (let i = startRow; i < lastRow; i++) {
                const cell = sheet['F' + i]
                if (cell !== undefined) {
                  const pages = Number(cell.v)
                  sum += pages
                  let date = sheet['C' + i].v.toString()
                  date += sheet['D' + i].v.toString()
                  date += sheet['E' + i].v.toString()
                  const doc = {
                    title: sheet['B' + i].v,
                    date,
                    pages 
                  }
                  docs.push(doc)
                  count += 1
                }
              }

              // console.log(count)
              // console.log(sum)
              // console.log(docs)

              let arr0 = []
              arr0[1] = name
              arr0[6] = count
              arr0[7] = sum
              ws_data0.push(arr0)

              let ws_data1 = [
                [ "序号", "案卷号", "案卷级档号", "档号", "类号", "类别代号", "类别件号", '材料名称', '形成时间', '页数', '' ]
              ]
              let arr1 = []
              arr1[0] = name
              ws_data1.push(arr1)

              for (let i = 1; i <= docs.length; i++) {
                const doc = docs[i-1]
                arr1 = []
                arr1[0] = i
                arr1[7] = doc.title
                arr1[8] = doc.date
                arr1[9] = doc.pages
                ws_data1.push(arr1)
              }

              const wb = utils.book_new()
              utils.book_append_sheet(wb, utils.aoa_to_sheet(ws_data1), sheetName)
              const data = write(wb, { type: "buffer", bookType: "xlsx" })
              await writeBinaryFile(newDir + '/' + name + '-人事卷内目录.xlsx', data)

            } catch(err) {
              console.log(err)
            }
          } else {
            // if entry is dir
          }
        }

        const wb = utils.book_new()
        utils.book_append_sheet(wb, utils.aoa_to_sheet(ws_data0), sheetName)
        const data = write(wb, { type: "buffer", bookType: "xlsx" })
        await writeBinaryFile(newDir + '/人事案卷.xlsx', data)

        setMsg('完成')

        // open newDir
      } catch(err) {
        console.log(err)
        setMsg('只能选择目录')
      }
    }
  }

  document.addEventListener('contextmenu', event => event.preventDefault());

  return (
    <div className="container">
      <div className="row">
        <a href="https://itove.com" target="_blank">
          <img src="/tauri.svg" className="logo tauri" alt="itove logo" />
        </a>
      </div>

      <h4>人事档案转科怡</h4>
      <ul className="tip">
      <li>将人事档案表格转换为科怡支持表格格式</li>
      <li>选择人事档案表格所在目录，点击<strong>开始转换</strong></li>
      <li>转换后的表格将保存在同目录下</li>
      </ul>

      <div className="row">
        <form className="form"
          onSubmit={(e) => {
            e.preventDefault();
            main()
          }}
        >
          <label>人事档案文件夹<span className="asteroid">*</span></label>
          <input
            className="input"
            onClick={() => getDir()}
            readOnly
            required
            placeholder="点击选择目录或将目录拖拽到这里"
            value={dir}
          />
          <label>人员基本信息表</label>
          <input
            className="input"
            onClick={() => getFile()}
            readOnly
            required
            placeholder="点击选择人员基本信息表"
            value={file}
          />
          <button type="submit" className="btn">开始转换</button>
        </form>
      </div>
      <p>{msg}</p>
      <p className="footer">{appName} {ver} <br/>
      更新日志
      </p>

    </div>
  );
}

export default App;
