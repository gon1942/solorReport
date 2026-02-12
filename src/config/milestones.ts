export type MilestoneTemplate = {
  id: string;
  title: string;
  description: string;
};

export const milestoneTemplates: Record<string, MilestoneTemplate[]> = {
  capex: [
    {
      id: 'm1',
      title: '투자 승인 준비',
      description: 'CAPEX 승인에 필요한 자료와 산출물을 준비합니다.'
    },
    {
      id: 'm2',
      title: '재무 검토',
      description: '재무 모델과 투자 타당성을 검토합니다.'
    },
    {
      id: 'm3',
      title: '최종 의사결정',
      description: '의사결정 위원회에 상정하고 최종 결정을 확정합니다.'
    }
  ]
};
