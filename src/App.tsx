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
  const [prefix, setPrefix] = useState("")
  const sheetName = "Sheet1"

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
    if (file !== null) {
      console.log(file)
      setFile(file)
    }
  }

  function isXlsx() {
    return false
  }

  async function importXlsx(file) {
    let data = {}
    return data
  }

  async function exportAnJuan(aoa) {
    const ws_data = [
      ...[[ "案卷级档案", "姓名", "性别", "身份证号", "政治面貌", "密集架号", "总件数", '总页数' ]],
      ...aoa
    ]
    console.log(ws_data);
    const wb = utils.book_new()
    utils.book_append_sheet(wb, utils.aoa_to_sheet(ws_data), sheetName)
    const data = write(wb, { type: "buffer", bookType: "xlsx" })
    await writeBinaryFile(dir + '/' + outputDir + '/人事案卷.xlsx', data)
  }

  async function extractMoreData(name, sheet) {
    for (const key in sheet) {
      if (sheet[key].v === name) {
        const row = key.replace(/[A-Z]*/, '')

        const gender = sheet['C' + row].v
        const id = sheet['D' + row].v
        const org = sheet['E' + row].v
        const birth = sheet['F' + row].v
        const ethnic = sheet['G' + row].v
        const homeland = sheet['H' + row].v
        const status = sheet['K' + row].v
        return {
          name,
          gender,
          id,
          org,
          birth,
          ethnic,
          homeland,
          status
        }
      }
    }
  }

  async function extractData(sheet) {
    const name = sheet.A2.v.split('：')[2].replace(/ /g, '')
    const sn = sheet.A2.v.split('：')[1].replace(/[^\x00-\x7F]/g, "").replace(/ /g, '')

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
      const cell = sheet['G' + i]
      if (cell !== undefined) {
        const pages = Number(cell.v)
        sum += pages
        let date
        if (sheet['C' + i] !== undefined) {
          date = sheet['C' + i].v.toString()
          date += sheet['D' + i].v.toString()
          date += sheet['E' + i].v.toString()
        }
        const doc = {
          sn,
          cateid: sheet['A' + i].v,
          title: sheet['B' + i].v,
          date,
          pages 
        }
        docs.push(doc)
        count += 1
      }
    }

    return {
      name,
      sn,
      count,
      sum,
      docs
    }
  }

  async function exportJuanNei(individual) {
    let ws_data = [
      [ "序号", "案卷号", "案卷级档号", "档号", "类号", "类别代号", "类别件号", '材料名称', '形成时间', '页数', '' ]
    ]
    let aoa = []
    aoa[0] = individual.name
    ws_data.push(aoa)

    for (let i = 1; i <= individual.docs.length; i++) {
      const doc = individual.docs[i-1]
      aoa = []
      aoa[0] = i
      aoa[7] = doc.title
      aoa[8] = doc.date
      aoa[9] = doc.pages
      ws_data.push(aoa)
    }

    const wb = utils.book_new()
    utils.book_append_sheet(wb, utils.aoa_to_sheet(ws_data), sheetName)
    const contents = write(wb, { type: "buffer", bookType: "xlsx" })
    await writeBinaryFile(dir + '/' + outputDir + '/' + individual.name + '-人事卷内目录.xlsx', contents)
  }

  async function getSheet(path) {
    const contents = await readBinaryFile(path)
    const workbook = read(contents)
    const sheeName = workbook.SheetNames[0]
    return workbook.Sheets[sheeName]
  }

  async function main() {
    if (file === null || file === '') {
    } else {
      const basicInfoSheet = await getSheet(file)
      // const basicInfo = await extractMoreData('李志浩', basicInfoSheet)
      // console.log(basicInfo)
    }

    if (dir === null || dir === '') {
      setMsg('请选择目录')
    } else {
      try {
        const entries = await readDir(dir, { recursive: true })
        setMsg('处理中...')

        const newDir = dir + '/' + outputDir
        await createDir(newDir, { recursive: true })

        let aoa = []

        for (const entry of entries) {
          if (entry.children === undefined) {
            // if entry is file
            const individual = await extractData(await getSheet(entry.path))
            console.log(individual)
            // append AnJuan ws_data
            // aoa = [[1,2,3,4,5],[2,3,4,5,6]]
            
            // export JuanNei
            exportJuanNei(individual)
            let file = entry
          }
        }

        // export AnJuan
        exportAnJuan(aoa)

        setMsg('完成')

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
          <label>案卷级档号前缀<span className="asteroid">*</span><br/>格式：YCSY01-RST01-1</label>
          <input
            name="prefix"
            className="input"
            required
            placeholder="填写案卷级档号前缀"
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
