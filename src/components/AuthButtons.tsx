import React from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../api/authConfig";

export const AuthButtons: React.FC = () => {
  const { instance, accounts } = useMsal();
  const isAuthenticated = accounts.length > 0;

  const handleLogin = () => {
    instance.loginRedirect(loginRequest).catch((e) => {
      console.error("로그인 실패:", e);
    });
  };

  const handleLogout = () => {
    instance.logoutRedirect().catch((e) => {
      console.error("로그아웃 실패:", e);
    });
  };

  if (isAuthenticated) {
    return (
      <div className="flex items-center justify-between p-4 glass-panel rounded-2xl animate-fade-in">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-full bg-gradient-to-tr from-brand-500 to-purple-500 flex items-center justify-center text-white font-bold shadow-lg">
            {accounts[0].username.charAt(0).toUpperCase()}
          </div>
          <div className="text-sm">
            <p className="text-slate-400 text-xs uppercase tracking-wider font-semibold">연동된 계정</p>
            <p className="font-medium text-slate-100">{accounts[0].username}</p>
          </div>
        </div>
        <button
          onClick={handleLogout}
          className="px-4 py-2 text-sm font-medium text-slate-300 bg-white/5 hover:bg-white/10 hover:text-white rounded-xl transition-all duration-200 border border-white/5 hover:border-white/10"
        >
          로그아웃
        </button>
      </div>
    );
  }

  return (
    <div className="p-8 glass-panel rounded-2xl text-center animate-slide-up flex flex-col items-center justify-center min-h-[300px]">
      <div className="w-16 h-16 mb-6 rounded-2xl bg-gradient-to-br from-brand-500 to-purple-600 flex items-center justify-center shadow-lg shadow-brand-500/20">
        <svg className="w-8 h-8 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
        </svg>
      </div>
      <h3 className="mb-3 text-2xl font-bold text-white tracking-tight">
        Microsoft 계정 연동
      </h3>
      <p className="mb-8 text-slate-400 max-w-sm mx-auto leading-relaxed">
        복잡한 메일 분류 규칙을 서버와 동기화하기 위해 Outlook 계정 인증이 필요합니다.
      </p>
      <button
        onClick={handleLogin}
        className="group relative inline-flex items-center justify-center px-8 py-3.5 text-base font-semibold text-white transition-all duration-200 bg-brand-600 border border-transparent rounded-xl hover:bg-brand-500 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-brand-600 focus:ring-offset-slate-900 overflow-hidden"
      >
        <span className="absolute inset-0 w-full h-full -mt-1 rounded-lg opacity-30 bg-gradient-to-b from-transparent via-transparent to-black"></span>
        <span className="relative flex items-center gap-2">
          Microsoft 계정으로 시작하기
          <svg className="w-5 h-5 transition-transform duration-200 group-hover:translate-x-1" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M14 5l7 7m0 0l-7 7m7-7H3" />
          </svg>
        </span>
      </button>
    </div>
  );
};
