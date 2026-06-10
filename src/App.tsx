import React from 'react';
import { MsalProvider } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { AuthButtons } from './components/AuthButtons';
import { RuleBuilder } from './components/RuleBuilder';

interface AppProps {
  msalInstance: PublicClientApplication;
}

function App({ msalInstance }: AppProps) {
  // 컴포넌트 마운트 시 msalInstance에서 계정 정보를 직접 가져옴
  const accounts = msalInstance.getAllAccounts();
  const isAuthenticated = accounts.length > 0;

  return (
    <MsalProvider instance={msalInstance}>
      <div className="min-h-screen p-6 md:p-12 lg:p-24 flex flex-col items-center bg-gray-950 text-slate-200">
        
        {/* 헤더 섹션 */}
        <header className="w-full max-w-4xl text-center mb-12 animate-slide-up">
          <div className="inline-flex items-center gap-2 px-3 py-1 mb-6 text-xs font-medium text-brand-400 bg-brand-500/10 rounded-full border border-brand-500/20">
            <span className="w-2 h-2 rounded-full bg-brand-500 animate-pulse"></span>
            Chrome Extension Preview
          </div>
          <h1 className="text-5xl md:text-6xl font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-white via-slate-200 to-slate-400 mb-6 tracking-tight">
            Outlook Smart Router
          </h1>
          <p className="text-lg text-slate-400 max-w-2xl mx-auto font-light leading-relaxed">
            한계를 뛰어넘는 메일 자동 분류 규칙. <br className="hidden md:block" />
            <strong className="font-medium text-slate-200">AND/OR 조건</strong>을 자유롭게 결합하여 나만의 완벽한 메일 정리 환경을 구축하세요.
          </p>
        </header>

        {/* 메인 컨텐츠 영역 */}
        <main className="w-full max-w-4xl grid gap-8 animate-fade-in" style={{ animationDelay: '0.2s' }}>
          <AuthButtons />
          
          {/* 사용자가 인증되었을 때만 규칙 빌더를 보여줍니다 */}
          {isAuthenticated && (
            <RuleBuilder />
          )}
        </main>
        
      </div>
    </MsalProvider>
  );
}

export default App;
