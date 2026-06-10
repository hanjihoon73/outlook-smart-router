import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { PublicClientApplication } from "@azure/msal-browser"
import { msalConfig } from "./api/authConfig"
import './index.css'
import App from './App.tsx'

const msalInstance = new PublicClientApplication(msalConfig);

msalInstance.initialize().then(() => {
  // 인증 후 돌아왔을 때 토큰을 처리하는 필수 로직
  msalInstance.handleRedirectPromise().then(() => {
    createRoot(document.getElementById('root')!).render(
      <StrictMode>
        <App msalInstance={msalInstance} />
      </StrictMode>,
    )
  }).catch((e) => {
    console.error("인증 리디렉트 처리 실패:", e);
  });
});
