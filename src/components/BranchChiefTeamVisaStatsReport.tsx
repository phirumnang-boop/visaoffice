import React from 'react';
import {
  TeamVisaIssuanceStatsReport,
  TeamVisaIssuanceStatsReportProps,
} from './TeamVisaIssuanceStatsReport';

export const BranchChiefTeamVisaStatsReport: React.FC<TeamVisaIssuanceStatsReportProps> = (props) => {
  return <TeamVisaIssuanceStatsReport {...props} isBranchChiefMode={true} />;
};

export default BranchChiefTeamVisaStatsReport;
