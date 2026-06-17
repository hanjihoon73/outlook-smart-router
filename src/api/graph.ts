import { InteractionRequiredAuthError, type IPublicClientApplication } from "@azure/msal-browser";
import { loginRequest } from "./authConfig";

const GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0";

/**
 * 활성화된 계정을 통해 MS Graph API 액세스 토큰을 획득합니다.
 * 토큰이 만료되었거나 추가 인증이 필요한 경우 팝업을 띄워 다시 인증을 요청합니다.
 */
export const getAccessToken = async (msalInstance: IPublicClientApplication): Promise<string> => {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) {
    throw new Error("활성화된 계정이 없습니다. 다시 로그인해 주세요.");
  }
  
  const request = {
    ...loginRequest,
    account: accounts[0]
  };

  try {
    // 조용히(백그라운드에서) 토큰 갱신 시도
    const response = await msalInstance.acquireTokenSilent(request);
    return response.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      // 백그라운드 갱신 실패 시 팝업으로 재인증
      const response = await msalInstance.acquireTokenPopup(request);
      return response.accessToken;
    }
    throw error;
  }
};

/**
 * 사용자의 메일 폴더 목록을 조회합니다.
 */
export const getMailFolders = async (accessToken: string) => {
  const headers = new Headers();
  headers.append("Authorization", `Bearer ${accessToken}`);

  const options = {
    method: "GET",
    headers: headers
  };

  // 최대 50개의 상위 폴더를 가져옵니다.
  const response = await fetch(`${GRAPH_ENDPOINT}/me/mailFolders?$top=50`, options);
  if (!response.ok) {
    throw new Error("폴더 목록을 불러오는 데 실패했습니다.");
  }
  return response.json();
};

/**
 * 받은 편지함(inbox)에 새로운 메일 분류 규칙을 생성합니다.
 */
export const createMessageRule = async (accessToken: string, rulePayload: any) => {
  const headers = new Headers();
  headers.append("Authorization", `Bearer ${accessToken}`);
  headers.append("Content-Type", "application/json");

  const options = {
    method: "POST",
    headers: headers,
    body: JSON.stringify(rulePayload)
  };

  const response = await fetch(`${GRAPH_ENDPOINT}/me/mailFolders/inbox/messageRules`, options);
  if (!response.ok) {
    throw new Error("규칙 생성에 실패했습니다.");
  }
  return response.json();
};

/**
 * 기존 메일 분류 규칙을 업데이트합니다. (On/Off 토글 등에 사용)
 */
export const updateMessageRule = async (accessToken: string, ruleId: string, rulePayload: any) => {
  const headers = new Headers();
  headers.append("Authorization", `Bearer ${accessToken}`);
  headers.append("Content-Type", "application/json");

  const options = {
    method: "PATCH",
    headers: headers,
    body: JSON.stringify(rulePayload)
  };

  const response = await fetch(`${GRAPH_ENDPOINT}/me/mailFolders/inbox/messageRules/${ruleId}`, options);
  if (!response.ok) {
    throw new Error("규칙 업데이트에 실패했습니다.");
  }
  return response.json();
};

/**
 * 기존 메일 분류 규칙을 삭제합니다.
 */
export const deleteMessageRule = async (accessToken: string, ruleId: string) => {
  const headers = new Headers();
  headers.append("Authorization", `Bearer ${accessToken}`);

  const options = {
    method: "DELETE",
    headers: headers
  };

  const response = await fetch(`${GRAPH_ENDPOINT}/me/mailFolders/inbox/messageRules/${ruleId}`, options);
  if (!response.ok) {
    throw new Error(`규칙 삭제에 실패했습니다. (ID: ${ruleId})`);
  }
  return true;
};
