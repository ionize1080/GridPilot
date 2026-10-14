const {contextBridge,ipcRenderer}=require('electron');
contextBridge.exposeInMainWorld('gridpilot',{request:message=>ipcRenderer.invoke('request',message),onProgress:callback=>{const listener=(_event,data)=>callback(data);ipcRenderer.on('progress',listener);return ()=>ipcRenderer.removeListener('progress',listener);},close:()=>ipcRenderer.send('close')});
