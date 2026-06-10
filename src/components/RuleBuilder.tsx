import React, { useState, useEffect } from 'react';
import { useMsal } from "@azure/msal-react";
import { getAccessToken, getMailFolders, createMessageRule } from '../api/graph';
import { compileToGraphRules } from '../utils/compiler';

// DNF 방식의 그룹 관리를 위한 데이터 구조
type Field = 'from' | 'subject' | 'body';
type Operator = 'contains' | 'equals';

interface Condition {
  id: string;
  field: Field;
  operator: Operator;
  value: string;
}

interface ConditionGroup {
  id: string;
  conditions: Condition[]; // 이 그룹 내의 조건들은 AND로 묶임
}

export const RuleBuilder: React.FC = () => {
  const { instance } = useMsal();
  const [folders, setFolders] = useState<any[]>([]);
  const [loadingFolders, setLoadingFolders] = useState<boolean>(true);
  const [folderError, setFolderError] = useState<string>('');
  const [isSaving, setIsSaving] = useState<boolean>(false);

  const [groups, setGroups] = useState<ConditionGroup[]>([
    {
      id: 'group-1',
      conditions: [{ id: 'cond-1', field: 'from', operator: 'contains', value: '' }]
    }
  ]);

  // 폴더 이동 액션 상태
  const [targetFolder, setTargetFolder] = useState<string>('');

  // 컴포넌트 마운트 시 실제 사용자의 메일 폴더 목록 불러오기
  useEffect(() => {
    const fetchFolders = async () => {
      try {
        setFolderError('');
        const token = await getAccessToken(instance);
        const data = await getMailFolders(token);
        setFolders(data.value || []);
      } catch (error: any) {
        console.error("폴더 목록 불러오기 실패:", error);
        setFolderError(error.message || "권한이 없거나 불러오기에 실패했습니다.");
      } finally {
        setLoadingFolders(false);
      }
    };
    fetchFolders();
  }, [instance]);

  const addGroup = () => {
    setGroups([
      ...groups,
      {
        id: `group-${Date.now()}`,
        conditions: [{ id: `cond-${Date.now()}`, field: 'from', operator: 'contains', value: '' }]
      }
    ]);
  };

  const addCondition = (groupId: string) => {
    setGroups(groups.map(g => {
      if (g.id === groupId) {
        return {
          ...g,
          conditions: [...g.conditions, { id: `cond-${Date.now()}`, field: 'from', operator: 'contains', value: '' }]
        };
      }
      return g;
    }));
  };

  const updateCondition = (groupId: string, condId: string, updates: Partial<Condition>) => {
    setGroups(groups.map(g => {
      if (g.id === groupId) {
        return {
          ...g,
          conditions: g.conditions.map(c => c.id === condId ? { ...c, ...updates } : c)
        };
      }
      return g;
    }));
  };

  const removeGroup = (groupId: string) => {
    setGroups(groups.filter(g => g.id !== groupId));
  };

  const removeCondition = (groupId: string, condId: string) => {
    setGroups(groups.map(g => {
      if (g.id === groupId) {
        return {
          ...g,
          conditions: g.conditions.filter(c => c.id !== condId)
        };
      }
      return g;
    }));
  };

  const handleSaveRule = async () => {
    try {
      setIsSaving(true);
      // 1. UI 상태를 MS Graph API가 이해할 수 있는 순수 AND 규칙 배열로 변환
      const graphRules = compileToGraphRules({ groups, targetFolder });
      
      if (graphRules.length === 0) {
        alert("저장할 유효한 규칙이 없습니다. 조건을 입력해 주세요.");
        return;
      }

      // 2. 액세스 토큰 획득
      const token = await getAccessToken(instance);

      // 3. 변환된 모든 규칙을 서버로 전송
      for (const rulePayload of graphRules) {
        await createMessageRule(token, rulePayload);
      }

      alert("🎉 메일 분류 규칙이 성공적으로 동기화되었습니다!");
      
    } catch (error: any) {
      console.error("규칙 저장 실패:", error);
      alert(`오류가 발생했습니다: ${error.message}`);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="w-full glass-panel rounded-2xl p-6 md:p-8 animate-fade-in text-left">
      <div className="mb-8 border-b border-white/10 pb-4">
        <h2 className="text-2xl font-bold text-white tracking-tight flex items-center gap-2">
          <svg className="w-6 h-6 text-brand-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
          </svg>
          고급 메일 분류 규칙 만들기
        </h2>
        <p className="text-sm text-slate-400 mt-2">
          각 박스 안의 조건들은 <strong>'그리고(AND)'</strong>로 결합되며, 박스와 박스 사이는 <strong>'또는(OR)'</strong>로 결합됩니다.
        </p>
      </div>

      <div className="space-y-6">
        {groups.map((group, groupIndex) => (
          <div key={group.id} className="relative">
            {/* OR 구분자 (첫 번째 그룹 제외) */}
            {groupIndex > 0 && (
              <div className="absolute -top-4 left-1/2 -translate-x-1/2 -translate-y-1/2 z-10">
                <div className="bg-slate-800 text-brand-400 text-xs font-bold px-3 py-1 rounded-full border border-slate-700 shadow-md uppercase tracking-widest">
                  OR (또는)
                </div>
              </div>
            )}

            {/* AND 그룹 박스 (글래스모피즘) */}
            <div className={`p-6 rounded-2xl border ${groupIndex === 0 ? 'border-brand-500/30 bg-brand-500/5' : 'border-purple-500/30 bg-purple-500/5'} transition-all duration-300 hover:border-white/20`}>
              <div className="flex flex-col gap-3">
                {/* 그룹 삭제 버튼 (첫 번째 그룹은 삭제 불가) */}
                {groups.length > 1 && (
                  <button 
                    onClick={() => removeGroup(group.id)}
                    className="absolute top-2 right-2 text-slate-500 hover:text-red-400 p-1 transition-colors"
                    title="이 그룹 삭제"
                  >
                    <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                    </svg>
                  </button>
                )}

                {group.conditions.map((cond, condIndex) => (
                  <div key={cond.id} className="flex items-center gap-3">
                    {/* 조건 블록 UI */}
                    <div className="flex-1 flex flex-wrap md:flex-nowrap items-center gap-2 bg-slate-800/80 p-2 pl-4 rounded-xl border border-white/5 shadow-inner">
                      {condIndex > 0 && <span className="text-xs font-bold text-slate-500 mr-2">AND</span>}
                      
                      <select 
                        value={cond.field}
                        onChange={(e) => updateCondition(group.id, cond.id, { field: e.target.value as Field })}
                        className="bg-slate-900 border border-slate-700 text-sm text-slate-200 rounded-lg px-3 py-2 focus:ring-2 focus:ring-brand-500 outline-none transition-all"
                      >
                        <option value="from">보낸 사람</option>
                        <option value="subject">제목</option>
                        <option value="body">본문</option>
                      </select>
                      
                      <select 
                        value={cond.operator}
                        onChange={(e) => updateCondition(group.id, cond.id, { operator: e.target.value as Operator })}
                        className="bg-slate-900 border border-slate-700 text-sm text-slate-200 rounded-lg px-3 py-2 focus:ring-2 focus:ring-brand-500 outline-none transition-all"
                      >
                        <option value="contains">포함 (Contains)</option>
                        <option value="equals">일치 (Equals)</option>
                      </select>

                      <input 
                        type="text" 
                        placeholder="검색어 입력..."
                        value={cond.value}
                        onChange={(e) => updateCondition(group.id, cond.id, { value: e.target.value })}
                        className="flex-1 bg-slate-900 border border-slate-700 text-sm text-slate-100 rounded-lg px-4 py-2 focus:ring-2 focus:ring-brand-500 outline-none transition-all"
                      />

                      {/* 조건 삭제 버튼 */}
                      {group.conditions.length > 1 && (
                        <button 
                          onClick={() => removeCondition(group.id, cond.id)}
                          className="ml-2 p-1.5 text-slate-500 hover:text-red-400 hover:bg-red-400/10 rounded-lg transition-colors"
                          title="조건 삭제"
                        >
                          <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                          </svg>
                        </button>
                      )}
                    </div>
                  </div>
                ))}
              </div>
              
              {/* AND 조건 추가 버튼 */}
              <button 
                onClick={() => addCondition(group.id)}
                className="mt-4 text-xs font-medium text-brand-400 hover:text-brand-300 flex items-center gap-1 px-3 py-1.5 rounded-lg hover:bg-brand-500/10 transition-colors"
              >
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
                </svg>
                AND 조건 추가
              </button>
            </div>
          </div>
        ))}
      </div>

      {/* OR 그룹 추가 버튼 */}
      <div className="mt-8 flex justify-center border-t border-dashed border-white/10 pt-6">
        <button 
          onClick={addGroup}
          className="group relative inline-flex items-center justify-center px-6 py-2.5 text-sm font-semibold text-white transition-all duration-200 bg-slate-800 border border-slate-700 rounded-xl hover:bg-slate-700 hover:border-slate-500 shadow-lg"
        >
          <span className="flex items-center gap-2">
            <svg className="w-4 h-4 text-brand-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6v6m0 0v6m0-6h6m-6 0H6" />
            </svg>
            새로운 조건 그룹 추가 (OR)
          </span>
        </button>
      </div>

      {/* 실행할 액션 (폴더 이동) 섹션 */}
      <div className="mt-12 pt-8 border-t-2 border-white/5 animate-slide-up" style={{ animationDelay: '0.1s' }}>
        <h3 className="text-xl font-bold text-white mb-6 flex items-center gap-2">
          <svg className="w-6 h-6 text-emerald-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
          </svg>
          실행할 액션
        </h3>
        
        <div className="p-6 rounded-2xl border border-emerald-500/30 bg-emerald-500/5 shadow-inner">
          <div className="flex flex-wrap md:flex-nowrap items-center gap-4">
            <div className="flex items-center gap-2 text-emerald-300 font-medium whitespace-nowrap bg-emerald-500/10 px-4 py-2 rounded-lg">
              <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z" />
              </svg>
              다음 폴더로 이동:
            </div>
            
            <select 
              value={targetFolder}
              onChange={(e) => setTargetFolder(e.target.value)}
              className="flex-1 bg-slate-900 border border-slate-700 text-slate-200 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-emerald-500 outline-none transition-all"
            >
              <option value="" disabled>
                {loadingFolders 
                  ? "폴더 목록 불러오는 중..." 
                  : folderError 
                    ? `오류 발생 (클릭 불가)` 
                    : "이동할 폴더를 선택하세요..."}
              </option>
              {folders.map((folder) => (
                <option key={folder.id} value={folder.id}>
                  {folder.displayName}
                </option>
              ))}
            </select>
          </div>
          {folderError && (
            <div className="mt-3 text-sm text-red-400 bg-red-400/10 px-4 py-2 rounded-lg border border-red-400/20">
              폴더 로드 오류: {folderError}
              <br/>새로고침하여 다시 로그인 승인을 받아주세요.
            </div>
          )}
        </div>
      </div>

      {/* 저장 버튼 */}
      <div className="mt-10 flex justify-end">
        <button 
          onClick={handleSaveRule}
          disabled={isSaving || !targetFolder}
          className={`px-8 py-3.5 ${isSaving || !targetFolder ? 'bg-slate-700 cursor-not-allowed opacity-50' : 'bg-gradient-to-r from-brand-600 to-purple-600 hover:from-brand-500 hover:to-purple-500'} text-white font-bold rounded-xl shadow-lg shadow-brand-500/20 transition-all duration-300 flex items-center gap-2 group`}
        >
          {isSaving ? (
            <span className="flex items-center gap-2">
              <svg className="animate-spin -ml-1 mr-2 h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
              </svg>
              저장 중...
            </span>
          ) : (
            <>
              <svg className="w-5 h-5 transition-transform group-hover:scale-110" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4" />
              </svg>
              규칙 저장 및 동기화
            </>
          )}
        </button>
      </div>

    </div>
  );
};
