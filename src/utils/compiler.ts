export type Field = 'from' | 'subject' | 'body' | 'to' | 'cc';
export type Operator = 'contains' | 'equals';

export interface Condition {
  id: string;
  field: Field;
  operator: Operator;
  value: string;
}

export interface ConditionGroup {
  id: string;
  conditions: Condition[];
}

export interface RuleBuilderState {
  groups: ConditionGroup[];
  targetFolder: string;
}

/**
 * 사용자 UI의 상태를 Microsoft Graph API의 messageRule 페이로드 배열로 변환합니다.
 * Outlook 기본 규칙은 'AND' 조건만 지원하므로,
 * UI의 여러 'OR' 그룹(ConditionGroup)들은 각각 독립된 여러 개의 규칙(Rule)으로 쪼개서 생성합니다.
 */
export const compileToGraphRules = (state: RuleBuilderState, ruleNamePrefix: string = "SmartRouter_") => {
  if (!state.targetFolder) {
    throw new Error("이동할 대상 폴더가 선택되지 않았습니다.");
  }

  const graphRules: any[] = [];

  state.groups.forEach((group, index) => {
    // 빈 조건 필터링
    const validConditions = group.conditions.filter(c => c.value.trim() !== '');
    if (validConditions.length === 0) return; // 내용이 없는 그룹은 무시

    const conditionsPayload: any = {};

    validConditions.forEach(cond => {
      const val = cond.value.trim();
      
      switch (cond.field) {
        case 'from':
          if (cond.operator === 'contains') {
            conditionsPayload.senderContains = conditionsPayload.senderContains 
              ? [...conditionsPayload.senderContains, val] 
              : [val];
          } else if (cond.operator === 'equals') {
            const address = { emailAddress: { address: val } };
            conditionsPayload.fromAddresses = conditionsPayload.fromAddresses 
              ? [...conditionsPayload.fromAddresses, address] 
              : [address];
          }
          break;
          
        case 'subject':
          // Graph API의 subjectContains는 기본적으로 부분 일치를 지원
          conditionsPayload.subjectContains = conditionsPayload.subjectContains 
            ? [...conditionsPayload.subjectContains, val] 
            : [val];
          break;
          
        case 'body':
          conditionsPayload.bodyContains = conditionsPayload.bodyContains 
            ? [...conditionsPayload.bodyContains, val] 
            : [val];
          break;
          
        case 'to':
        case 'cc':
          // MS Graph API는 받는 사람과 참조 문자열 매칭을 recipientContains로 통합 검색 지원합니다.
          conditionsPayload.recipientContains = conditionsPayload.recipientContains 
            ? [...conditionsPayload.recipientContains, val] 
            : [val];
          break;
      }
    });

    // 해당 그룹을 위한 단일 Graph API 규칙 조각 생성
    const rulePayload = {
      displayName: `${ruleNamePrefix}Group_${index + 1}`,
      sequence: 10 + index, // 실행 우선순위 (임의 지정)
      isEnabled: true,
      conditions: conditionsPayload,
      actions: {
        moveToFolder: state.targetFolder
      }
    };

    graphRules.push(rulePayload);
  });

  return graphRules;
};
