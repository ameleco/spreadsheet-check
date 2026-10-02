let initialized=null, checking=false;
const base='https://cdn.jsdelivr.net/pyodide/v314.0.7/full/';
async function initialize(id) {
  self.postMessage({id,stage:'Loading the spreadsheet checker…'});
  const {loadPyodide}=await import(base+'pyodide.mjs');
  const py=await loadPyodide({indexURL:base,stdout:()=>{},stderr:()=>{}});
  await py.loadPackage('micropip');
  self.postMessage({id,stage:'Preparing lab rubrics…'});
  await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl==3.1.5')");
  const response=await fetch(new URL('./grader.py',import.meta.url));if(!response.ok)throw Error('Rubrics unavailable');
  py.runPython(await response.text());
  const adapter=await fetch(new URL('./student_adapter.py',import.meta.url));if(!adapter.ok)throw Error('File checks unavailable');
  py.runPython(await adapter.text());
  return py;
}
self.onmessage=async({data})=>{
  if(checking)return;checking=true;const {id,lab,buffer}=data;let py;
  try {
    initialized??=initialize(id);py=await initialized;
    self.postMessage({id,stage:'Checking your tables, calculations, and charts…'});
    py.FS.writeFile('/tmp/student.xlsx',new Uint8Array(buffer));
    py.globals.set('_selected_lab',lab);
    const result=JSON.parse(py.runPython("student_check('/tmp/student.xlsx', _selected_lab)"));
    if(result.error)self.postMessage({id,error:result.error});else self.postMessage({id,result});
  } catch {initialized=null;self.postMessage({id,error:'The checker couldn’t load or read this spreadsheet. Check your connection and try again. If it continues, ask your instructor to review the file.'});}
  finally {if(py){try{py.FS.unlink('/tmp/student.xlsx');py.globals.delete('_selected_lab');}catch{}}checking=false;}
};
