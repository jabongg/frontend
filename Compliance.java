package com.swte.insurance.reports.editor;

import java.util.Calendar;

import org.apache.log4j.Logger;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.jface.dialogs.TrayDialog;
import org.eclipse.nebula.widgets.calendarcombo.CalendarCombo;
import org.eclipse.swt.SWT;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Control;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.ui.IWorkbenchWindow;

import com.swte.SWTEException;
import com.swte.common.ui.control.SwteIntegerControl;
import com.swte.common.ui.control.ValidatorHelper;
import com.swte.common.ui.factory.FactoryUI;
import com.swte.insurance.finance.ReportMessage;
import com.swte.util.CommonUtil;

public class ActivePolicyComplianceReportDialog extends TrayDialog 
{
     private static Logger logger = Logger.getLogger(ActivePolicyComplianceReportDialog.class);
     private static final String POLICY_ID="Policy Id";
     private static final String ACTIVE_POLICY_FROM="Active Policy From";
     private static final String ACTIVE_POLICY_TO="Active Policy To";
     private static final int NUM_OF_COLUMNS=2,LIMIT=10;
     private static final String ACTIVE_POLICY_COMPLIANCE_REPORT="Active Policy Compliance Report";
     private SwteIntegerControl policyIdtxt;
     protected CalendarCombo activePolicyFrom,activePolicyTo;
     protected Button exportExcelButton;
     private static final int PRINT_REPORT= 0;
     private static final String EXPORT_EXCEL ="Export to Excel";
     private static final String CLOSE_BUTTON ="Close";
     public ActivePolicyComplianceReportDialog(Shell parentShell, IWorkbenchWindow window) {
          super(parentShell);
     }
     public Control createDialogArea(Composite shell)
     {
          Composite comp = (Composite) super.createDialogArea(shell);
          GridLayout layout = (GridLayout) comp.getLayout();
          layout.numColumns = NUM_OF_COLUMNS;
          comp.getShell().setText(ACTIVE_POLICY_COMPLIANCE_REPORT);
          FactoryUI.createLabel(comp,POLICY_ID);
          policyIdtxt=FactoryUI.createIntegerTextField(comp, true, LIMIT);
          FactoryUI.createLabel(comp,ACTIVE_POLICY_FROM);
          activePolicyFrom=FactoryUI.createCalendarComboWithFormat(comp, SWT.None);
          FactoryUI.createLabel(comp, ACTIVE_POLICY_TO);
          activePolicyTo=FactoryUI.createCalendarComboWithFormat(comp, SWT.None);
          initialzeDefaultValue();
          return comp;
     }
     private void initialzeDefaultValue() {
          activePolicyFrom.setDate(CommonUtil.getPreviousDayTimestamp());
          activePolicyTo.setDate(CommonUtil.getCurrentTimestamp());
     }
     public Point getInitialSize(){
          return new Point(300,200);
     }
     @Override
     protected void createButtonsForButtonBar(Composite parent) {
          exportExcelButton = createButton(parent, PRINT_REPORT, EXPORT_EXCEL, false);
          //exportExcelButton.forceFocus();
          createButton(parent, CANCEL,CLOSE_BUTTON, false);
     }
     protected String getReportEditorName() {
          return null;
     }
     @Override
     protected void buttonPressed(int buttonId) {
          switch (buttonId) {
          case PRINT_REPORT:
              try {
					  if(!validateInputFileds()) return;
					  printComplianceReport();
              } catch (SWTEException e) {
                   e.printStackTrace();
                   logger.error(e.getMessage());
              }
              break;

          default:
              close();
              break;
          }
     }
     private boolean validateInputFileds(){

          String policyId  = policyIdtxt.getText();
          Calendar activePolicyDateFrom=activePolicyFrom.getDate();
          Calendar activePolicyDateTo=activePolicyTo.getDate();
          int i=activePolicyDateTo.compareTo(activePolicyDateFrom);
          if(policyId!=null && !policyId.isEmpty()) {
              boolean isPolicyIdValid = ValidatorHelper.isLong(policyId);
          
              if(!isPolicyIdValid) {
                   MessageDialog.openError(getShell(), ReportMessage.ERROR, "Policy id should be numeric")  ;
                   return false;
              }
          }
          
          if(i==-1)
          {
              MessageDialog.openError(getShell(), ReportMessage.ERROR, "activePolicyFrom smaller than activePolicyTo");
              return false;
          }
          if(activePolicyFrom.getDate() == null){
              MessageDialog.openError(getShell(), ReportMessage.MISSING_INPUT, ReportMessage.ADD_PARA_FROMDATE);
              return false;
          }
          if(activePolicyTo.getDate() == null){
              MessageDialog.openError(getShell(), ReportMessage.MISSING_INPUT, ReportMessage.ADD_PARA_TODATE);
              return false;
          } 
              return true;

     }
          
 	public void printComplianceReport() throws DocumentException, IOException,SWTEException {

		// From Date
		
		Calendar fromDateCal = activePolicyDate.getDate();
		Timestamp fromDateTimestamp = null;
		if(fromDateCal != null){
			fromDateTimestamp = new Timestamp(fromDateCal.getTimeInMillis());
		}
		ArrayList<Object> inputParaList = new ArrayList<Object>();
		inputParaList.add(fromDateTimestamp);
		inputParaList.add(fromDateTimestamp);
		String SHEET_NAME = "Compliance Report";
		
		List<ReportingBean> appTmgBeansList = ServiceLocator.getReportService().getActivePolicyList(inputParaList);
		String nameStr = FinancialReportUtil.getReportName("ActivePolicyReport", fromDateTimestamp, fromDateTimestamp);

		String filePath  = FinancialReportUtil.reportPath(nameStr);
		ExcelReportWriter excelReportWriter = new ExcelReportWriter();
		excelReportWriter.setReportPoperties(SHEET_NAME, filePath);

		excelReportWriter.writeReport(appTmgBeansList, activePolicyHeaderLabels(),activePolicyColumnList());
		FinancialReportUtil.openFile(filePath);
	}
	         
     
     protected String getReportName() {
          // TODO Auto-generated method stub
          return null;
     }
}

