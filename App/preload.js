const {contextBridge,ipcRenderer}=require('electron');
contextBridge.exposeInMainWorld('gridpilot',{request:message=>ipcRenderer.invoke('request',message),close:()=>ipcRenderer.send('close')});
