const DEFAULT_API='http://127.0.0.1:8000';
const input=document.querySelector('#api');
const status=document.querySelector('#status');
chrome.storage.local.get('apiUrl').then(({apiUrl})=>{input.value=apiUrl||DEFAULT_API});
document.querySelector('#settings').addEventListener('submit',async event=>{
  event.preventDefault();
  try{
    const url=new URL(input.value.trim());
    if(!['http:','https:'].includes(url.protocol)||url.username||url.password)throw new Error('Enter an http or https URL without credentials.');
    // Chrome match patterns use an optional port; omitting it grants this host
    // across ports and avoids treating URL.origin (which includes the port) as
    // a permission pattern.
    const hostPattern=`${url.protocol}//${url.hostname}/*`;
    if(url.origin!==new URL(DEFAULT_API).origin&&!await chrome.permissions.request({origins:[hostPattern]})){
      throw new Error(`Chrome denied access to ${hostPattern}. Grant this host permission to connect to the backend; the URL was not saved.`);
    }
    await chrome.storage.local.set({apiUrl:url.origin});
    status.textContent='Backend URL saved.';
  }catch(error){status.textContent=error.message||'Could not save backend URL.';}
});
