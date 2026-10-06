import React from 'react';
import { EVisaRobokReport } from './EVisaStockReport';
import { CategoriesState, Officer, StockRecord } from '../types';

export interface EVisaTeamRobokReportProps {
  stockRecords: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  currentRole?: string;
  userName?: string;
  assignedTeam?: string;
  onClose?: () => void;
  onShowToast?: (message: string, type?: 'success' | 'error' | 'info') => void;
}

export const EVisaTeamRobokReport: React.FC<EVisaTeamRobokReportProps> = ({
  stockRecords,
  categories,
  officers,
  currentRole,
  assignedTeam,
  onClose,
}) => {
  const isTeamUser = currentRole !== 'Secondary' && currentRole !== 'Admin' && currentRole !== 'User (ការិយាល័យ)';

  return (
    <EVisaRobokReport
      stockRecords={stockRecords}
      categories={categories}
      officers={officers}
      onClose={onClose}
      isSingleTeam={true}
      hideSwitcher={true}
      assignedTeam={assignedTeam}
      lockTeamSelection={Boolean(assignedTeam)}
      reportTitle="របកក្រដាសអនុម័តក្រុម"
    />
  );
};

export default EVisaTeamRobokReport;
