// UI feedback helpers: toasts and async confirm dialog
// Designed to replace alert()/confirm() without blocking the UI.

const TOAST_CONTAINER_ID = 'toastContainer';
const CONFIRM_DIALOG_ID = 'dlgConfirm';

function ensureToastContainer(){
  let c = document.getElementById(TOAST_CONTAINER_ID);
  if(!c){
    c = document.createElement('div');
    c.id = TOAST_CONTAINER_ID;
    c.className = 'toast-container';
    document.body.appendChild(c);
  }
  return c;
}

/**
 * Shows a toast message.
 * @param {string} message
 * @param {'info'|'success'|'warning'|'error'} [type='info']
 * @param {number} [timeoutMs=3500]
 */
export function toast(message, type='info', timeoutMs=3500){
  const cont = ensureToastContainer();
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.innerHTML = `
    <div class="toast-body">${escapeHtml(String(message || ''))}</div>
    <button class="toast-close" aria-label="Fechar">×</button>
  `;
  const close = () => {
    el.classList.add('toast-hide');
    setTimeout(() => el.remove(), 180);
  };
  el.querySelector('.toast-close')?.addEventListener('click', close);
  cont.appendChild(el);
  // trigger animation
  requestAnimationFrame(() => el.classList.add('toast-show'));
  if(timeoutMs > 0){
    setTimeout(close, timeoutMs);
  }
}

function ensureConfirmDialog(){
  let dlg = document.getElementById(CONFIRM_DIALOG_ID);
  if(!dlg){
    dlg = document.createElement('dialog');
    dlg.id = CONFIRM_DIALOG_ID;
    dlg.className = 'dialog';
    dlg.innerHTML = `
      <form method="dialog" class="form">
        <h3 id="dlgConfirmTitle">Confirmar</h3>
        <div id="dlgConfirmMsg" class="muted" style="margin-top:6px;white-space:pre-wrap"></div>
        <menu style="margin-top:14px">
          <button value="cancel">Cancelar</button>
          <button value="ok" class="primary">Confirmar</button>
        </menu>
      </form>
    `;
    document.body.appendChild(dlg);
  }
  return dlg;
}

/**
 * Async confirm (dialog). Usage: if(await confirmDlg('...')) { ... }
 * @param {string} message
 * @param {{title?:string, okText?:string, cancelText?:string, danger?:boolean}} [opts]
 * @returns {Promise<boolean>}
 */
export function confirmDlg(message, opts={}){
  const dlg = ensureConfirmDialog();
  const title = dlg.querySelector('#dlgConfirmTitle');
  const msg = dlg.querySelector('#dlgConfirmMsg');
  const btnOk = dlg.querySelector('button.primary');
  const btnCancel = dlg.querySelector('button:not(.primary)');

  if(title) title.textContent = opts.title || 'Confirmar';
  if(msg) msg.textContent = String(message || '');
  if(btnOk) {
    btnOk.textContent = opts.okText || 'Confirmar';
    btnOk.classList.toggle('danger', !!opts.danger);
  }
  if(btnCancel) btnCancel.textContent = opts.cancelText || 'Cancelar';

  return new Promise(resolve => {
    const onClose = () => {
      dlg.removeEventListener('close', onClose);
      const v = dlg.returnValue;
      resolve(v === 'ok');
    };
    dlg.addEventListener('close', onClose);
    dlg.showModal();
  });
}

export function escapeHtml(s){
  return s
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'",'&#39;');
}
