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
const outputDir = '转科怡'
interface Entries {
  children?: Entries[],
  name: string,
  path: string
}

function App() {
  const [msg, setMsg] = useState("")
  const [msg2, setMsg2] = useState("")
  const [msg3, setMsg3] = useState("")
  const [dir, setDir] = useState("")
  const [basicInfoFile, setBasicInfoFile] = useState("")
  const [archiveSnFile, setArchiveSnFile] = useState("")
  const sheetName = "Sheet1"
  let files: Entries[] = []
  let basicInfoSheet: any
  let archiveSnSheet: any

  listen<string>('tauri://file-drop', (event) => {
    setDir(event.payload[0])
  })

  async function getDir() {
    const dir = await open({directory: true})
    if (!Array.isArray(dir) && dir !== null) {
      setDir(dir)
    }
  }

  async function getBasicInfoFile() {
    const file = await open({directory: false})
    if (!Array.isArray(file) && file !== null && file !== '') {
      setBasicInfoFile(file)
    }
  }

  async function getArchiveSnFile() {
    const file = await open({directory: false})
    if (!Array.isArray(file) && file !== null && file !== '') {
      setArchiveSnFile(file)
    }
  }

  function isXlsx(filePath: string) {
    const ext = filePath.split('.')[filePath.split.length - 1]
    if (ext === 'xlsx' || ext === 'xls') {
      return true
    } else {
      return false
    }
  }

  async function exportAnJuan(aoa: any) {
    const ws_data = [
      ...[[ "案卷级档案", "姓名", "性别", "身份证号", "政治面貌", "密集架号", "总件数", '总页数', '单位' ]],
      ...aoa
    ]
    const wb = utils.book_new()
    utils.book_append_sheet(wb, utils.aoa_to_sheet(ws_data), sheetName)
    const data = write(wb, { type: "buffer", bookType: "xlsx" })
    await writeBinaryFile(dir + '/' + outputDir + '/人事案卷.xlsx', data)
  }

  async function extractMoreData(name: string, sheet: any) {
    for (const key in sheet) {
      if (sheet[key].v === name) {
        const row = key.replace(/[A-Z]*/, '')

        const gender = sheet['C' + row].v
        const id = sheet['D' + row].v
        const org = sheet['E' + row].v
        const birth = sheet['F' + row].v
        const ethnic = sheet['G' + row].v
        const homeland = sheet['H' + row].v
        const status = sheet['K' + row] !== undefined ? sheet['K' + row].v : ''
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

  async function extractData(sheet: any) {
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
        }
        if (sheet['D' + i] !== undefined) {
          date += sheet['D' + i].v.toString().padStart(2, 0)
        }
        if (sheet['E' + i] !== undefined) {
          date += sheet['E' + i].v.toString().padStart(2, 0)
        }
        const doc = {
          sn,
          cateid: sheet['A' + i] !== undefined ? sheet['A' + i].v : '',
          title: sheet['B' + i] !== undefined ? sheet['B' + i].v : '',
          date,
          pages 
        }
        docs.push(doc)
        count += 1
      }
    }

    let info = {
      name,
      sn,
      count,
      sum,
      docs,
      archiveSn: '',
      more: {
        gender: '',
        id: '',
        status: '',
        org: ''
      }
    }

    if (basicInfoSheet !== null && basicInfoSheet !== undefined) {
      const basicInfo = await extractMoreData(name, basicInfoSheet)
      if (basicInfo !== undefined ) {
        info.more = basicInfo
      }
    } 

    if (archiveSnSheet !== null && archiveSnSheet !== undefined) {
      const archiveSn = await extractArchiveSn(info.more.org, archiveSnSheet)
      if (archiveSn !== undefined ) {
        info.archiveSn = archiveSn
        // console.log(archiveSn)
      }
    } 

    return info
  }

  async function extractArchiveSn(org: string, sheet: any) {
    for (const key in sheet) {
      if (sheet[key].v === org) {
        const row = key.replace(/[A-Z]*/, '')
        const sn = sheet['C' + row].v
        return sn.substr(0, sn.lastIndexOf('-'))
      }
    }
  }

  async function exportJuanNei(individual: any) {
    let ws_data = [
      [ "序号", "案卷号", "案卷级档号", "档号", "类号", "类别代号", "类别件号", '材料名称', '形成时间', '页数', '' ]
    ]
    let aoa = []
    aoa[0] = individual.name
    ws_data.push(aoa)

    for (let i = 1; i <= individual.docs.length; i++) {
      const doc = individual.docs[i-1]
      aoa = []
      aoa[0] = i  //序号
      aoa[1] = individual.sn  //案卷号
      aoa[2] = individual.archiveSn + '-' + individual.sn //案卷级档号
      aoa[3] = individual.archiveSn + '-' + individual.sn + '-' + i //档号
      aoa[4] = doc.cateid //类别号
      aoa[5] = doc.cateid.split('-')[0] //类别代号
      aoa[6] = doc.cateid.split('-')[doc.cateid.split('-').length - 1] //类别件号
      aoa[7] = doc.title //材料名称
      aoa[8] = doc.date //形成时间
      aoa[9] = doc.pages //页数
      ws_data.push(aoa)
    }

    const wb = utils.book_new()
    utils.book_append_sheet(wb, utils.aoa_to_sheet(ws_data), sheetName)
    const contents = write(wb, { type: "buffer", bookType: "xlsx" })
    await writeBinaryFile(dir + '/' + outputDir + '/' + individual.name + '-人事卷内目录.xlsx', contents)
  }

  async function getSheet(path: string) {
    const contents = await readBinaryFile(path)
    const workbook = read(contents)
    const sheeName = workbook.SheetNames[0]
    return workbook.Sheets[sheeName]
  }

  async function getFilesInDir(entries: any[]) {
    for (const e of entries) {
      if (e.children === undefined) {
        if (isXlsx(e.path)) {
          files.push(e)
        }
      } else if (e.name !== outputDir) {
        getFilesInDir(e.children)
      }
    }
  }

  async function main() {
    if (dir === null || dir === '') {
      setMsg('请选择目录')
    } else {
      try {
        const entries = await readDir(dir, { recursive: true })
        setMsg('处理中...')

        const newDir = dir + '/' + outputDir
        await createDir(newDir, { recursive: true })

        let aoa = []

        await getFilesInDir(entries)

        if (basicInfoFile !== null && basicInfoFile !== '') {
          basicInfoSheet = await getSheet(basicInfoFile)
        }

        if (archiveSnFile !== null && archiveSnFile !== '') {
          archiveSnSheet = await getSheet(archiveSnFile)
        }

        for (const file of files) {
          setMsg2(file.path)
          const individual = await extractData(await getSheet(file.path))
          let arr = []
          arr[0] = individual.archiveSn + '-' + individual.sn //案卷级档号
          arr[1] = individual.name //姓名
          if (individual.more !== undefined) {
            arr[2] = individual.more.gender //性别
            arr[3] = individual.more.id //身份证号
            arr[4] = individual.more.status //政治面貌
          }
          // arr[5] = '' //密集架号
          arr[6] = individual.count //总件数
          arr[7] = individual.sum //总页数
          arr[8] = individual.more.org //单位
          aoa.push(arr)

          // export JuanNei
          exportJuanNei(individual)
        }

        // export AnJuan
        exportAnJuan(aoa)

        setMsg2(`输出目录：[${dir}/${outputDir}/]`)
        setMsg('完成')

      } catch(err) {
        console.log(err)
        // setMsg('只能选择目录')
        setMsg('表格格式不匹配')
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
      <li>选择人事档案表格所在目录</li>
      <li>填写案卷级档号前缀</li>
      <li>选择人员基本信息表（提供性别、身份证号、政治面貌等信息，非必须）</li>
      <li>选择单位案卷级档号对应表（提供单位案卷级档号，非必须）</li>
      <li>点击<strong>开始转换</strong></li>
      <li>转换后的表格将保存在同目录下的<strong>转科怡</strong>目录</li>
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
            onClick={() => getBasicInfoFile()}
            readOnly
            required
            placeholder="点击选择人员基本信息表"
            value={basicInfoFile}
          />
          <label>单位案卷级档号对应表</label>
          <input
            className="input"
            onClick={() => getArchiveSnFile()}
            readOnly
            required
            placeholder="点击选择单位案卷级档号对应表"
            value={archiveSnFile}
          />
          <button type="submit" className="btn">开始转换</button>
        </form>
      </div>
      <p>{msg}</p>
      <p>{msg2}</p>
      <p className="footer">{appName} {ver} <br/>
      更新日志
      </p>

    </div>
  );
}

export default App;
