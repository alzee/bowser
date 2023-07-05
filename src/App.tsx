import { useState } from "react";
import reactLogo from "./assets/react.svg";
import { invoke } from "@tauri-apps/api/tauri";
import "./App.css";
import { read, utils, writeFile } from 'xlsx';
import { useEffect } from 'react';
import { open } from '@tauri-apps/api/dialog';
import { renameFile, readDir } from '@tauri-apps/api/fs';
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

  async function go(dir: string, digit: number) {
    /**
    if (dir === null || dir === '') {
      setMsg('请选择目录')
    } else {
      try {
        const entries = await readDir(dir)
        setMsg('处理中...')
        for (const entry of entries) {
          // ignore dirs
          if (entry.children === undefined) {
            const arr = entry.name!.split('.')
            const ext = arr.pop()
            // remove leadding 0s so don't need to worry about name.length gt or lt digit
            const name = arr.join('.').replace(/^0+/, '')
            const newName = name.padStart(digit, '0')
            await renameFile(entry.path, dir + '/' + newName + '.' + ext);
          }
        }
        setMsg('完成')
      } catch(err) {
        setMsg('只能选择目录');
      }
    }
   **/
  }

  document.addEventListener('contextmenu', event => event.preventDefault());

  return (
    <div className="container">
      <div className="row">
        <a href="https://itove.com" target="_blank">
          <img src="/tauri.svg" className="logo tauri" alt="itove logo" />
        </a>
      </div>

      <h4>修改文件名长度</h4>
      <ul className="tip">
      <li>当文件名大于指定长度时，删除文件名中多余的前缀 0 </li>
      <li>当文件名小于指定长度时，用 0 前置填充文件名</li>
      </ul>

      <div className="row">
        <form
          onSubmit={(e) => {
            e.preventDefault();
            go(dir, digit)
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
