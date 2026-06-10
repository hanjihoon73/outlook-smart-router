import type { Configuration, PopupRequest } from "@azure/msal-browser";

// 주의: 아래 clientId는 사용자가 Azure Portal에서 생성한 앱의 클라이언트 ID로 변경해야 합니다.
export const msalConfig: Configuration = {
  auth: {
    clientId: "4c1d847a-6bb9-4499-bba8-280228517a85", // TODO: Azure 앱 등록 후 발급받은 Client ID 입력
    authority: "https://login.microsoftonline.com/5e8ceabe-d9ba-48bc-ab7e-e74bb9d2506d", // 단일 테넌트용 ID 적용
    redirectUri: "/", // 인증 후 돌아올 주소
  },
  cache: {
    cacheLocation: "sessionStorage", // 'localStorage' 또는 'sessionStorage'
    storeAuthStateInCookie: false, // IE11이나 구형 Edge를 위한 설정 (일반적으로 false)
  },
};

// 로그인 시 요청할 권한 (Scopes) 설정
export const loginRequest: PopupRequest = {
  scopes: ["User.Read", "MailboxSettings.ReadWrite", "Mail.Read"],
};
