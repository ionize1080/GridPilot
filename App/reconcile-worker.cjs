const {parentPort,workerData}=require('node:worker_threads');
try{parentPort.postMessage(require('./reconcile-core.cjs').reconcile(workerData.config,workerData.tables));}catch(e){parentPort.postMessage({error:e.message});}
