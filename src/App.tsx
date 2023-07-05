import { useState } from "react";
import reactLogo from "./assets/react.svg";
import { invoke } from "@tauri-apps/api/tauri";
import "./App.css";
import { read, readFile, utils, writeFile } from 'xlsx';
import { useEffect } from 'react';
import { open } from '@tauri-apps/api/dialog';
import { readDir, readBinaryFile, BaseDirectory } from '@tauri-apps/api/fs';
import { getName, getVersion } from '@tauri-apps/api/app';
import { listen } from '@tauri-apps/api/event';

const ver = await getVersion();
const appName = await getName();

function App() {
  const [name, setName] = useState("");
  const [digit, setDigit] = useState(2);
  const [msg, setMsg] = useState("");
  const [dir, setDir] = useState("");

  listen<string>('tauri://file-drop', (event) => {
    setDir(event.payload[0])
  })

  async function getDir() {
    const dir = await open({directory: true})
    if (!Array.isArray(dir) && dir !== null) {
      setDir(dir)
    }
  }

  async function go() {
    if (dir === null || dir === '') {
      setMsg('请选择目录')
    } else {
      try {
        const entries = await readDir(dir)
        setMsg('处理中...')
        for (const entry of entries) {
          // ignore dirs
          if (entry.children === undefined) {
            // console.log(entry)
            try {
              const contents = await readBinaryFile(dir + '/' + entry.name)
              // const x = read(dir + '/' + entry.name)
              const workbook = read(contents)
              console.log(workbook)
              // console.log(x.Sheets)
              // console.log(x.Sheets[sheet_name])
            } catch(err) {
              console.log(err)
            }
            // do the thing
            // is xlsx?
            // is format valid?
            // extract data
            // write to xlsx1
            // write to xlsx2
          }
        }
        setMsg('完成')
      } catch(err) {
        setMsg('只能选择目录');
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

      <h4>组织部转科怡</h4>
      <ul className="tip">
      <li>将组织部人事档案表格转换为科怡支持表格格式</li>
      <li>选择组织部人事档案表格所在目录，点击确定</li>
      <li>转换后的表格将保存在同目录下</li>
      </ul>

      <div className="row">
        <form
          onSubmit={(e) => {
            e.preventDefault();
            go()
          }}
        >
          <label>点击选择目录或将目录拖拽到这里</label>
          <input
            id="dir-input"
            onClick={() => getDir()}
            readOnly
            required
            placeholder="点击选择目录或将目录拖拽到这里"
            value={dir}
          />
          <div className="row2">
          <input
            id="digit"
            type="number"
            min="2"
            required
            onChange={(e) => setDigit(Number(e.currentTarget.value))}
            placeholder="文件名长度"
          />
          <button type="submit">确定</button>
          </div>
        </form>
      </div>
      <p>{msg}</p>
      <p className="footer">{appName} {ver}</p>

    </div>
  );
}

export default App;
