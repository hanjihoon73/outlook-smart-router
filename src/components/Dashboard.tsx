import React, { useState, useEffect } from 'react';
import { useMsal } from "@azure/msal-react";
import { supabase } from '../api/supabase';
import { getAccessToken, updateMessageRule, deleteMessageRule } from '../api/graph';
import { RuleBuilder } from './RuleBuilder';

export const Dashboard: React.FC = () => {
  const { instance, accounts } = useMsal();
  const userEmail = accounts[0]?.username;
  
  const [rules, setRules] = useState<any[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedRuleId, setSelectedRuleId] = useState<string | null>(null);
  const [newRuleKey, setNewRuleKey] = useState<number>(Date.now());

  useEffect(() => {
    if (userEmail) {
      fetchRules();
    }
  }, [userEmail]);

  const fetchRules = async () => {
    setLoading(true);
    const { data, error } = await supabase
      .from('mail_rules')
      .select('*')
      .eq('user_email', userEmail)
      .order('created_at', { ascending: false });
      
    if (error) {
      console.error("규칙 목록 불러오기 실패:", error);
    } else if (data) {
      setRules(data);
    }
    setLoading(false);
  };

  const handleToggleActive = async (ruleId: string, currentStatus: boolean, graphRuleIds?: string[]) => {
    const newStatus = !currentStatus;
    
    // 1. Supabase 상태 업데이트
    const { error } = await supabase
      .from('mail_rules')
      .update({ is_active: newStatus })
      .eq('id', ruleId);
      
    if (error) {
      alert('데이터베이스 상태 업데이트에 실패했습니다.');
      return;
    }

    // 2. MS Graph API 동기화
    if (graphRuleIds && graphRuleIds.length > 0) {
      try {
        const token = await getAccessToken(instance);
        for (const graphId of graphRuleIds) {
          await updateMessageRule(token, graphId, { isEnabled: newStatus });
        }
      } catch (err) {
        console.error('MS 서버 규칙 동기화 실패:', err);
        alert('MS 서버와 동기화하는 도중 일부 오류가 발생했습니다.');
      }
    }

    fetchRules();
  };

  const handleDelete = async (ruleId: string, graphRuleIds?: string[]) => {
    if (!window.confirm("이 규칙을 완전히 삭제하시겠습니까?")) return;
    
    // 1. MS Graph 서버에서 규칙 삭제
    if (graphRuleIds && graphRuleIds.length > 0) {
      try {
        const token = await getAccessToken(instance);
        for (const graphId of graphRuleIds) {
          try {
            await deleteMessageRule(token, graphId);
          } catch (e) {
            console.warn(`MS 서버 규칙 삭제 실패 (이미 지워졌을 수 있음): ${graphId}`);
          }
        }
      } catch (err) {
        console.error('MS 서버 통신 오류:', err);
      }
    }

    // 2. Supabase DB에서 규칙 삭제
    const { error } = await supabase
      .from('mail_rules')
      .delete()
      .eq('id', ruleId);
      
    if (!error) {
      if (selectedRuleId === ruleId) setSelectedRuleId(null);
      fetchRules();
    } else {
      alert('데이터베이스에서 규칙을 삭제하는 데 실패했습니다.');
    }
  };

  const handleSaveSuccess = () => {
    fetchRules();
    if (!selectedRuleId) {
      // 새 규칙을 저장한 경우, 폼을 초기화하기 위해 key를 변경하여 리마운트시킵니다.
      setNewRuleKey(Date.now());
    }
  };

  const selectedRule = rules.find(r => r.id === selectedRuleId);

  return (
    <div className="flex flex-col lg:flex-row gap-6 w-full max-w-7xl mx-auto animate-fade-in">
      {/* 왼쪽: 규칙 리스트 (Master) */}
      <div className="w-full lg:w-1/3 glass-panel rounded-2xl p-6 flex flex-col h-auto lg:h-[calc(100vh-12rem)]">
        <div className="flex justify-between items-center mb-6">
          <h2 className="text-xl font-bold text-white">나의 분류 규칙</h2>
          <button 
            onClick={() => setSelectedRuleId(null)}
            className="p-2 bg-brand-500/20 text-brand-400 hover:bg-brand-500 hover:text-white rounded-lg transition-colors"
            title="새 규칙 만들기"
          >
            <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
            </svg>
          </button>
        </div>

        <div className="flex-1 overflow-y-auto space-y-3 pr-2 custom-scrollbar">
          {loading ? (
            <div className="text-slate-400 text-sm text-center py-4">불러오는 중...</div>
          ) : rules.length === 0 ? (
            <div className="text-slate-500 text-sm text-center py-8">
              저장된 규칙이 없습니다.<br/>우측 상단의 [+] 버튼을 눌러보세요.
            </div>
          ) : (
            rules.map(rule => (
              <div 
                key={rule.id}
                className={`p-4 rounded-xl border transition-all cursor-pointer group ${selectedRuleId === rule.id ? 'bg-brand-500/10 border-brand-500' : 'bg-slate-800/50 border-white/5 hover:border-brand-500/30'}`}
                onClick={() => setSelectedRuleId(rule.id)}
              >
                <div className="flex justify-between items-start mb-2">
                  <h3 className="text-slate-100 font-bold truncate pr-2">{rule.name}</h3>
                  <div className="flex items-center gap-2" onClick={e => e.stopPropagation()}>
                    {/* 토글 스위치 */}
                    <button 
                      onClick={() => handleToggleActive(rule.id, rule.is_active, rule.graph_rule_ids)}
                      className={`relative inline-flex h-5 w-9 items-center rounded-full transition-colors ${rule.is_active ? 'bg-brand-500' : 'bg-slate-600'}`}
                    >
                      <span className={`inline-block h-3 w-3 transform rounded-full bg-white transition-transform ${rule.is_active ? 'translate-x-5' : 'translate-x-1'}`} />
                    </button>
                    {/* 삭제 버튼 */}
                    <button 
                      onClick={() => handleDelete(rule.id, rule.graph_rule_ids)}
                      className="text-slate-500 hover:text-red-400 p-1 opacity-0 group-hover:opacity-100 transition-opacity"
                    >
                      <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                      </svg>
                    </button>
                  </div>
                </div>
                <div className="flex items-center gap-2 text-xs text-slate-400">
                  <svg className="w-3 h-3 text-emerald-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z" />
                  </svg>
                  <span className="truncate">{rule.target_folder_name || '폴더 지정 안됨'}</span>
                </div>
              </div>
            ))
          )}
        </div>
      </div>

      {/* 오른쪽: 상세 편집 (Detail) */}
      <div className="w-full lg:w-2/3 h-auto lg:h-[calc(100vh-12rem)] overflow-y-auto lg:pr-2 custom-scrollbar">
        <RuleBuilder 
          key={selectedRuleId || newRuleKey} 
          rule={selectedRule} 
          onSaveSuccess={handleSaveSuccess}
          userEmail={userEmail}
        />
      </div>
    </div>
  );
};
