package com.askey.agile.ex;


import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.ChangeConstants;
import com.agile.api.IAgileList;
import com.agile.api.IAgileSession;
import com.agile.api.ICell;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.INode;
import com.agile.api.IReferenceDesignatorCell;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.ICustomAction;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;
import com.agile.px.IWFChangeStatusEventInfo;
import com.askey.agile.appcode.AbeamUtil.AgileUtil;
import com.askey.agile.appcode.AbeamUtil.CommonUtil;
import com.askey.agile.appcode.AbeamUtil.LogUtils;
import com.askey.agile.appcode.agutils.PartsUtil;
import com.askey.agile.appcode.commonutils.ExceptionUtil;
import com.askey.agile.appcode.commonutils.Logutil;
import com.askey.agile.appcode.commonutils.MailUtil;
import com.askey.agile.constants.ASKEYBOMConstants;
import com.askey.agile.constants.ASKEYCommonConstants;
import com.askey.agile.constants.ASKEYECOConstants;
import com.askey.agile.constants.ASKEYERPConstants;
import com.askey.agile.constants.ASKEYPartsConstants;

import org.apache.logging.log4j.LogManager;  
import org.apache.logging.log4j.Logger;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.commons.collections.CollectionUtils;

/**
 *  目的：申請人送往下一個節點時,驗證bom是否存在錯誤,若存在錯誤退回申請人 
 *    01.	Find Number相關檢核
 *    02.	主替代料欄不得為空
 *    03.	主料不可重複
 *    04.	主料與替代料不可相同
 *    05.	替代料不可重複
 *    06.	主料的替代料代用方式必須為空
 *    07.	替代料代用方式不可為空
 *    08.	數量不得為空
 *    09.	替代料數量與主料數量需一致
 *    10.	若有插件位置,必須與數量一致
 *    11.	插件位置重複
 *    12.	替代料插件位置與主料插件位置需一致
 *    13.	68階製程不得重複且不得為空
 *    14.	非68階製程必須為空
 *    15.	替代料製程與主料製程需一致
 *    16.	當選到的是正常BOM,不允許料號為車載產品用料
 *    17.	插件位置長度超出15的限制  2020/9/4
 *    18.	替代料若有群組料且代用方式不為A及P,主料也必須為群組料 2024/1/31

 */


public class PC_15_01_BOMVerification implements IEventAction,ICustomAction{
	//private static Logger logger = null;
	private Logger logger = null;
	//private Logutil logger = new Logutil();
	private final String PROCESS_EXTENSION = "PC_15_01_BOMVerification";
	//private final String ID_PRIMARY_ITEM = "M";
	private final String VALUE_IS_Y = "Y"; 
	
	private final String ERROR_NULL_FINDNO = "\n▅👉Bom Item Find Number不為零或空或第1碼不為零:";
	private final String ERROR_THE_SAME_PRIMARY_SUBSTITUTE = "\n▅👉Find Number 主替料不可重覆:";
	private final String ERROR_DUPLICATE_SUBSTITUTE = "\n▅👉Find Number存在重複替代料:";
	private final String ERROR_DUPLICATE_FIND_NUMBER = "\n▅👉Find Number存在重複主料:";
	private final String ERROR_NULL_PRIMARY = "\n▅👉Find Number不存在主料:";
	
	private final String ERROR_REFSESIGNATOR_QTY = "\n▅👉料號存在插件位置與數量不一致:";
	private final String ERROR_NULL_PRIMARY_SUBSTITUE = "\n▅👉Bom Item主替代料欄位為空:";
	private final String ERROR_NULL_QTY = "\n▅👉Bom Item數量為空:";
	private final String ERROR_NULL_SUBSTITUTE_TYPE= "\n▅👉Bom Item替代料代用方式不可為空:";
	private final String ERROR_NULL_PRIMARY_TYPE= "\n▅👉Bom Item主料代用方式必須為空:";
	private final String ERROR_DUPLICATE_PRIMARY = "\n▅👉Bom Item存在重複主料:";
	private final String ERROR_DUPLICATE_BOARD_LAYRER = "\n▅👉Bom Item存在重複製程:";	
	private final String ERROR_NULL_BOARD_LAYRER = "\n▅👉Bom Item 製程欄位目前為空,不可為空:";
	private final String ERROR_NULL_ITEM = "\n▅👉Bom Item 不存在PLM系統:";
	private final String ERROR_NOT_NULL_BOARD_LAYRER = "\n▅👉Bom Item 製程欄位目前不為空,應為空:";
	private final String ERROR_PRELIMINARY_ITEM= "\n▅👉Bom Item 生命週期不得為Preliminary:";
	private final String ERROR_INCOSISENT_QTY= "\n▅👉Bom Item 替代料數量與主料不一致:";
	private final String ERROR_INCOSISENT_BOARD_LAYRER= "\n▅👉Bom Item 替代料製程（Board Layer）與主料不一致:";
	private final String ERROR_INCOSISENT_REFSESIGNATOR = "\n▅👉Bom Item 替代料插件位置與主料不一致:";	
	private final String WARNING_DUPLICATE_REFSESIGNATOR = "\n▅👉下面插件位置重複(警告):";
	//private final String WARNING_PROJECT_ITEM_OF_PART = "\nBom Item 不可為專案用料:(警告):";
	private final String WARNING_CAR_PLUGING_OF_PART = "\n▅👉Bom Item BOM中不允許加入車載產品用料 :(警告):";
	//private final String WARNING_DESIGN_CONTROL_OF_PART = "\nBom Item BOM中不允許加入設計管制料號 :(警告):";
	//private final String ERROR_68BOM_LAYER = "\n68Bom階  Item 製程欄位不允許為  CSP/COB/Boxbuild::";
	//private final String ERROR_70BOM_LAYER = "\n70Bom階  Item 製程欄位不允許為 Top/Bottom:";
	private final String ERROR_REF_OVER_MAX_LENGTH = String.format("\n▅👉插件位置超出ERP長度[%s]的限制:", ASKEYERPConstants.ERP_BOM_COMPONENT_REF_DESIGN_LENGTH);
	private final String ERROR_GROUPIPN_RULE_IS_VIOLATED = "\n▅👉違反群組料規則:替代料若為群組料且代用方式不為A及P,主料必為群組料:";
	private final String ERROR_HIGHKIGHT_PREFIX ="<br><FONT COLOR=red>";
	private final String HIGHKIGHT_SUFFIX ="</FONT>";
	private final int IDX_CHECK_BOM_ITEM=0;
	private final int IDX_CHECK_QTY=1;
	private final int IDX_CHECK_REFSESIGNATOR =2;
	private final int IDX_CHECK_BOARD_LAYRER=3;
	private final int IDX_CHECK_SUBSTITUTE_TYPE=4;
	private final int IDX_CHECK_IS_GROUP_PARTS=5;
	
	//private final int PROJECT_ITEM_OF_PART_BASEID = ASKEYPartsConstants.P2_PROJECT_ITEM_OF_PART;  						//專案用料 (Parts.P2  Base ID)
	//private final int PROJECT_ITEM_OF_AFFECTED_ITEMS_BASEID = ASKEYECOConstants.AFFECTED_ITEMS_PROJECT_ITEM_OF_PART;  	//專案用料 (Affected Items  Base ID)
	
	private final int CAR_PLUGING_OF_PART_BASEID = ASKEYPartsConstants.P2_IS_CAR_PLUGING; 							//是否為車載用料 (Parts.P2  Base ID)
	private final int CAR_PLUGING_OF_AFFECTED_ITEMS_BASEID = ASKEYECOConstants.AFFECTED_ITEMS_CAR_PLUGING_OF_PART;  //是否為車載用料 (Affected Items  Base ID)
	
	String errMsg = "";
	String formNumber = "";
	MailUtil smtpmail = null;
	
	public ActionResult doAction(IAgileSession session, INode actionNode, IDataObject affectedObject) {
		
		try {
			LogUtils.loadLog4j2Resource();
			logger = LogManager.getLogger(PC_15_01_BOMVerification.class);
			logger.info(LogUtils.formatHeading(PC_15_01_BOMVerification.class));		
			logger.info("-------------------------program start--------------------------");
			
			smtpmail = new MailUtil();			
			IChange change=(IChange) affectedObject;
			formNumber = change.getName();
			String errMsg = doIt(session, change,true);
			if (!"".equals(errMsg)) {
				logger.info("\n【BOM存在錯誤】\n" + errMsg);
				errMsg.replaceAll("\n", "<br>");				
				errMsg=ERROR_HIGHKIGHT_PREFIX+errMsg+HIGHKIGHT_SUFFIX;
			}
			else{
				errMsg="NO ERROR!";
				logger.info( "NO ERROR!" );
			}
			return new ActionResult(ActionResult.STRING, errMsg);
		} 
		catch (Exception ex) {
			ExceptionUtil exceptionUtil = new ExceptionUtil();			
			errMsg = "StackTrace(ex) : " + exceptionUtil.getExceptionInfo(ex, this.getClass().getName()) + "。" + "Message : " + ex.getMessage() + "。";
			logger.error(errMsg);
			smtpmail.OverrideSubject( formNumber + "-" + PROCESS_EXTENSION + " execute Fail ..." );
			smtpmail.sendMailByINI(errMsg , true);
			return new ActionResult(ActionResult.EXCEPTION, exceptionUtil.getExceptionInfoObject(ex, getClass()));			
		} 
		finally {
			logger.info("-------------------------program complete--------------------------");
		}
	}


	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo request) {
		try {
			LogUtils.loadLog4j2Resource();
			logger = LogManager.getLogger(PC_15_01_BOMVerification.class);
			logger.info(LogUtils.formatHeading(PC_15_01_BOMVerification.class));		
			logger.info("-------------------------program start--------------------------");

			smtpmail = new MailUtil();
			IWFChangeStatusEventInfo info = (IWFChangeStatusEventInfo) request;
			IChange change = (IChange) info.getDataObject();
			formNumber = change.getName();
			errMsg = "Fail to get Change: BOM Apply Change";
			errMsg = doIt(session, change,false);
			
			
			if (!"".equals(errMsg)) {
				logger.info("\n【BOM存在錯誤】\n" + errMsg);
				//logger.info("Show BOM Warning:" + errMsg);
				return new EventActionResult(request, new ActionResult(ActionResult.EXCEPTION, new Exception("\n"+errMsg) ));	
			}
			
			logger.info( "NO ERROR!" );
			return new EventActionResult(request,new ActionResult(ActionResult.STRING, "Run BOMVerifcation Successfully"));
		} 
		catch (Exception ex) {
			ExceptionUtil exceptionUtil = new ExceptionUtil();			
			errMsg = "StackTrace(ex) : " + exceptionUtil.getExceptionInfo(ex, this.getClass().getName()) + "。" + "Message : " + ex.getMessage() + "。";
			logger.error(errMsg);
			smtpmail.OverrideSubject( formNumber + "-" + PROCESS_EXTENSION + " execute Fail ..." );
			smtpmail.sendMailByINI(errMsg , true);
			return new EventActionResult(request, new ActionResult(ActionResult.EXCEPTION, exceptionUtil.getExceptionInfoObject(ex, getClass())));	
		}
		finally{
			logger.info("-------------------------program complete--------------------------");
		}
	}

	public String doIt(IAgileSession session, IChange change,boolean showWarning) throws Exception {
		//PC_21_01 會直接呼叫,所以這裡要先load log4j2
		LogUtils.loadLog4j2Resource();
		logger = LogManager.getLogger(PC_15_01_BOMVerification.class);
		logger.info(LogUtils.formatHeading(PC_15_01_BOMVerification.class));
		
		//IChange change = myChange;	
		PartsUtil partsutil = new PartsUtil();
		StringBuffer resultMsg = new StringBuffer();
		String RV_doIt = "";
		//ArrayList<String> arlAllowNormalBomUseProjectPart = new ArrayList<String>();
		ArrayList<String> arlAllowModelUsePartsOfDesignControlStatus = new ArrayList<String>();
		boolean existError; //判定bom是否存在錯誤
		boolean initBOMRelease = false; //是否首版BOM下發
		boolean rlAdded = false; //Redline標記為新增
		boolean rlModified = false; //Redline標記為修改
		boolean rlRemoved = false; //Redline標記為移除
		
		//專案用料 / 車載用料 / 設計管制料號
		//String projectItemOfBOM_YN = "";	//20210813 是否專案用料 與 允許非戰鬥機種使用專案用料的 欄位要取消 
		//String allowNormalBomUseProjectPartOfModel = ""; //20210813 是否專案用料 與 允許非戰鬥機種使用專案用料的 欄位要取消 
		
		String carPlugingOfBOM_YN = "";
		String bomModelName = "";
		String affNewRev = "";
		String affItemType = "";
		String itemCategory = "";		
		String allowModelUsePartsOfDesignControlStatus = "";	

		String findNo = "";
		String bomItemNo = "";
		String boardLayer = "";
		String primarySubstitute = "";
		String qty = "";
		String substituteType = "";
		String isGroupParts = "";

		String bomItemPrimarySubstitute = "";
		String bomItemFindNoSubstitute = "";
		String bomItemBoardLayer = "";
		String findNoPrimarySubstitute = "";
		String refDesigs="";
		String lifeCycle ="";
		String revision="";				
		String projectItemOfPart_YN = ""; //專案用料(子階)	
		
		final String initBOMRev = "E00"; //第一版BOM一定要是E00 (936 RULE)

		//驗證主替代為空的料號集合
		ArrayList<String> primarySubstituteEmptyList = null;
		//Preliminary料號集合
		ArrayList<String> preliminaryComList = null;
		ArrayList<String> preliminaryBOMList = null;
		//驗證位插件置數量不符料號集合
		ArrayList<String> refQtyErrorList = null;
		//驗證Find Number為空的料號集合
		ArrayList<String> findNoEmpList = null;
		//驗證不存在系統料號集合
		ArrayList<String> nullItemList = null;
		//驗證替代料代用方式為空的料號集合
		ArrayList<String> substituteTypeEmpList = null;
		//驗證主料代用方式不為空的料號集合
		ArrayList<String> substituteTypenNotEmpList = null;
		//驗證68製程為空的料號集合
		ArrayList<String> boardLayerEmpList = null;
		//驗證非68階製程不為空的料號集合
		ArrayList<String> boardLayerNotNullList = null;
		//驗證Find Number重複的集合
		ArrayList<String> duplicateFindNumberList = null;
		ArrayList<String> findNoList = null;
		//驗證數量為零或空的料號集合
		ArrayList<String> qtyZeroList = null;
		//驗證主料重複的集合
		ArrayList<String> duplicatePrimaryList = null;
		ArrayList<String> PrimaryList = null;
		//驗證重複替代料的集合
		ArrayList<String> duplicateSubstituteList = null;
		ArrayList<String> SubstituteList = null;
		//驗證製程重複的集合
		ArrayList<String> duplicateBoardLayerList = null;
		ArrayList<String> BoardLayerList = null;
		//驗證插件位置重複的集合
		ArrayList<String> duplicateRefDesigList = null;
		ArrayList<String> RefDesigListList = null;
		//驗證不存在主料
		HashMap<String,String[]> findNoMap = null;
		ArrayList<String> nullPrimaryList = null;
		//驗證主替代料不一致插件位置和數量
		HashMap<String,ArrayList<String[]>> findNoSubstituteMap = null;
		ArrayList<String> inconsisentQtyList = null;
		ArrayList<String> theSamePrimarySubstituteList = null;
		ArrayList<String> inconsisentRefDesigList = null;
		ArrayList<String> inconsisentBoardLayerList = null;	
		ArrayList<String> groupPartsRuleIsViolatedList = null;
		//不允許料號是否為 "專案用料"
		//ArrayList<String> ProjectItemOfPartList = null;
		//不允許料號為車載用料
		ArrayList<String>  CarPlugingOfPartList = null; 				
		//不允許設計管制料號入BOM
		//ArrayList<String> DesignControlOfPartList = null;
		//68Bom階  Item 製程欄位不允許為  CSP/COB/Boxbuild
		ArrayList<String> error68BOMLayerList = null;
		//70Bom階  Item 製程欄位不允許為 Top/Bottom";
		ArrayList<String> error70BOMLayerList = null;
		//插件位置超出長度限制(最大值: 15)
	    ArrayList<String> refOverMaxLengthList = null;
		
		try {
			logger.info("change:" + change);

			ITable affTbl = change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
			Iterator affIt = affTbl.iterator();			
			
			while (affIt.hasNext()) {
				//============= 重置變數===================
				//projectItemOfBOM_YN = "";				
				//allowNormalBomUseProjectPartOfModel = "";
				carPlugingOfBOM_YN = "";
				allowModelUsePartsOfDesignControlStatus = "";
				findNo = "";
				bomItemNo = "";
				boardLayer = "";
				primarySubstitute = "";
				qty = "";
				substituteType="";
				bomItemPrimarySubstitute = "";
				bomItemFindNoSubstitute = "";
				bomItemBoardLayer = "";
				findNoPrimarySubstitute = "";
				refDesigs="";
				lifeCycle ="";
				revision="";				
				projectItemOfPart_YN = ""; //專案用料(子階)
				
				//驗證主替代為空的料號集合
				primarySubstituteEmptyList = new ArrayList<String>();
				//Preliminary料號集合
				preliminaryComList = new ArrayList<String>();
				preliminaryBOMList = new ArrayList<String>();
				//驗證位插件置數量不符料號集合
				refQtyErrorList = new ArrayList<String>();
				//驗證Find Number為空的料號集合
				findNoEmpList = new ArrayList<String>();
				//驗證不存在系統料號集合
				nullItemList = new ArrayList<String>();
				//驗證替代料代用方式為空的料號集合
				substituteTypeEmpList = new ArrayList<String>();
				//驗證主料代用方式不為空的料號集合
				substituteTypenNotEmpList = new ArrayList<String>();
				//驗證68製程為空的料號集合
				boardLayerEmpList = new ArrayList<String>();
				//驗證非68階製程不為空的料號集合
				boardLayerNotNullList = new ArrayList<String>();
				//驗證Find Number重複的集合
				duplicateFindNumberList = new ArrayList<String>();
				findNoList = new ArrayList<String>();
				//驗證數量為零或空的料號集合
				qtyZeroList = new ArrayList<String>();
				//驗證主料重複的集合
				duplicatePrimaryList = new ArrayList<String>();
				PrimaryList = new ArrayList<String>();
				//驗證重複替代料的集合
				duplicateSubstituteList = new ArrayList<String>();
				SubstituteList = new ArrayList<String>();
				//驗證製程重複的集合
				duplicateBoardLayerList = new ArrayList<String>();
				BoardLayerList = new ArrayList<String>();
				//驗證插件位置重複的集合
				duplicateRefDesigList = new ArrayList<String>();
				RefDesigListList = new ArrayList<String>();
				//驗證不存在主料
				findNoMap=new HashMap<String,String[]> ();
				nullPrimaryList = new ArrayList<String>();
				//驗證主替代料不一致插件位置和數量
				findNoSubstituteMap=new HashMap<String,ArrayList<String[]>>();
				inconsisentQtyList= new ArrayList<String>();
				theSamePrimarySubstituteList= new ArrayList<String>();
				inconsisentRefDesigList = new ArrayList<String>();
				inconsisentBoardLayerList = new ArrayList<String>();
				groupPartsRuleIsViolatedList = new ArrayList<String>();
				//不允許料號是否為 "專案用料"
				//ProjectItemOfPartList = new ArrayList<String>();
				//不允許料號為車載用料
				CarPlugingOfPartList = new ArrayList<String>(); 				
				//不允許設計管制料號入BOM
				//DesignControlOfPartList = new ArrayList<String>();
				//68Bom階  Item 製程欄位不允許為  CSP/COB/Boxbuild (未上線)
				error68BOMLayerList = new ArrayList<String>();
				//70Bom階  Item 製程欄位不允許為 Top/Bottom (未上線)
				error70BOMLayerList = new ArrayList<String>();
				//插件位置超出長度限制(最大值: 15)
			    refOverMaxLengthList = new ArrayList<String>();
				//====================================================
				
				IRow affRow = (IRow) affIt.next();
				affNewRev = affRow.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).toString().toUpperCase();
				//initBOMRelease = ("".equals(affRow.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).toString())) ? true : false;
				if("".equals(affRow.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_OLD_REV).toString())) initBOMRelease = true; //old rev為空:為BOM首版
				if(initBOMRev.equals(affRow.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_NEW_REV).toString())) initBOMRelease = true; //new rev為E00:為BOM首版	
				
				//if(initBOMRelease)change.logAction("首發BOM");
				
				IItem affItem = (IItem) affRow.getReferent();
				
				//只驗証料號類
				if(! ASKEYPartsConstants.PART_SUPER_CLASS.equals(affItem.getAgileClass().getSuperClass().getName().toUpperCase()) ) continue;
				
				affItemType = affRow.getValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_TYPE).toString();
						
				//BOM 5899 及 1330 Subassembly的ModelName是放在P2,所以要判斷.
				if( affItemType.equals( ASKEYPartsConstants.SUB_CLASS_5899_SUB_PKG ) ||  affItemType.equals( ASKEYPartsConstants.SUB_CLASS_1330_SUBASSEMBLY_BUY_ITEM) ){
					bomModelName = affItem.getValue(ASKEYPartsConstants.P2_MODEL_NAME).toString();
				}
				else{
					bomModelName = affItem.getValue(ASKEYPartsConstants.TITLE_BLOCK_MODEL_NAME).toString();
				}
				
				//projectItemOfBOM_YN=affItem.getValue(new Integer(PROJECT_ITEM_OF_PART_BASEID)).toString().toUpperCase();
				carPlugingOfBOM_YN=affItem.getValue(new Integer(CAR_PLUGING_OF_PART_BASEID)).toString().toUpperCase();
				
				IRow rlP2Row = (IRow) affItem.getTable(ItemConstants.TABLE_REDLINEPAGETWO).iterator().next();
				/* 20210813 取消
				ICell cellProject_Item = rlP2Row.getCell(ASKEYPartsConstants.P2_PROJECT_ITEM_OF_PART);
				if(cellProject_Item == null){System.out.println("Project item is null ");}
				else{
					projectItemOfBOM_YN = cellProject_Item.getValue().toString().toUpperCase();
					System.out.println("Project item : " + projectItemOfBOM_YN);
				}
				*/
				ICell cellCarPluging_Item = rlP2Row.getCell(ASKEYPartsConstants.P2_IS_CAR_PLUGING);
				if(cellCarPluging_Item == null){System.out.println("Car pluging item is null ");}
				else{
					carPlugingOfBOM_YN = cellCarPluging_Item.getValue().toString().toUpperCase();
				}
				
				boolean is68BOM=affItem.getAgileClass().getName().startsWith("68")?true:false;
				boolean is70BOM=affItem.getAgileClass().getName().startsWith("70")?true:false;
				logger.info("\t\taffItem:" + affItem);
				logger.info("\t\tnewRelease:" + initBOMRelease);
				//errMsg.append(affItem.getName()).append(" BOM Verification Error:");
				ITable redlineBOM = affItem.getTable(ItemConstants.TABLE_REDLINEBOM);				
				Iterator rlIt = redlineBOM.iterator();
				existError = false;
				
				int refDesigQty=0;
				boolean existBOMError = false;
				if(!affItem.isFlagSet(ItemConstants.FLAG_HAS_RELEASED_REV))
					preliminaryBOMList.add(affItem.getName());
				ArrayList<String[]> substituteList=null;
				
				logger.info("\t\t\t BOM Check Start...");
				
				try{	
					while (rlIt.hasNext()) {
						IRow rlRow = (IRow) rlIt.next();
						lifeCycle = rlRow.getValue(ItemConstants.ATT_BOM_ITEM_LIFECYCLE_PHASE).toString();
						
						if (initBOMRelease) {
							//首版bom下發,視為所有修改標記均為新增
							//rlAdded = true;	
							/* */
							rlAdded = rlRow.isFlagSet(ItemConstants.FLAG_IS_REDLINE_ADDED);
							rlModified = rlRow.isFlagSet(ItemConstants.FLAG_IS_REDLINE_MODIFIED);
							rlRemoved = rlRow.isFlagSet(ItemConstants.FLAG_IS_REDLINE_REMOVED);
							
						} else {
							//確認Redline BOM的修改標記,只需檢查有變更的部份
							rlAdded = rlRow.isFlagSet(ItemConstants.FLAG_IS_REDLINE_ADDED);
							rlModified = rlRow.isFlagSet(ItemConstants.FLAG_IS_REDLINE_MODIFIED);
							rlRemoved = rlRow.isFlagSet(ItemConstants.FLAG_IS_REDLINE_REMOVED);
						}
						
						bomItemNo = rlRow.getValue(ItemConstants.ATT_BOM_ITEM_NUMBER).toString();
						boardLayer = rlRow.getValue(ASKEYPartsConstants.BOM_BOARD_LAYER).toString();
						primarySubstitute = rlRow.getValue(ASKEYPartsConstants.BOM_PRIMARY_SUBSTITUTE).toString();
						findNo = rlRow.getValue(ItemConstants.ATT_BOM_FIND_NUM).toString();
						isGroupParts = rlRow.getValue(ASKEYBOMConstants.BOM_IS_GROUP_PARTS).toString() ; //2023/11/8
						
						//2023/11/8: 判斷群組料規則,需先取得替代料的代用方式
						if (!"".equals(primarySubstitute) ) {
							if(ASKEYBOMConstants.BOM_PRIMARY_SUBSTITUTE_VALUEOF_S.equals( primarySubstitute )){
								substituteType = rlRow.getValue(ASKEYPartsConstants.BOM_SUBSTITUTE_TYPE).toString();
							}
						}

						//refDesigs = rlRow.getValue(ItemConstants.ATT_BOM_REF_DES).toString();	
						/* 20210622: 修改取得插件的值是自動展開後的值 */
						IReferenceDesignatorCell refCell = (IReferenceDesignatorCell) rlRow.getCell(ItemConstants.ATT_BOM_REF_DES);
						String refExStringValue = refCell.getExpandedValue();
						refDesigs = refExStringValue;
						
						if(refDesigs == null){
							refDesigs = "";
							//logger.info("\tRefDesign is Null, set value is string empty");
						}
						
						
						refDesigQty=0;
						qty=null;
						//logger.info("bomItemNo:"+bomItemNo+"\tfindNo:"+findNo+"\tM/S:"+primarySubstitute+"\tboardLayer:"+boardLayer);
						logger.debug("bomItemNumber = " + bomItemNo);
						
						//當選到的是正常BOM,不允許料號為專案用料
//						if(rlAdded || rlModified){							
//							if(! VALUE_IS_Y.equals(projectItemOfBOM_YN)){
//								IItem BomItem = (IItem) (IItem)session.getObject(IItem.OBJECT_TYPE, bomItemNo);
//								if( VALUE_IS_Y.equals( BomItem.getValue(new Integer(PROJECT_ITEM_OF_PART_BASEID) ).toString().toUpperCase() ) ){
//									ProjectItemOfPartList.add(bomItemNo);
//									existBOMError=true;
//								}
//							}							
//						}
						
						/* ******************************************************************************************
						 * 1).當選到的是正常BOM,不允許料號為專案用料 
						 * 2).ECN/DCN選到的是 “非戰鬥機種BOM” (正常BOM) ,若是零件為專案用料(Y)
						 *    且在料號的p2.允許非戰鬥機種使用專案用料  欄位的機種 與 改BOM的機種一樣,就可以加入此正常BOM
						 * 3).當選到的是正常BOM,不允許料號為車載產品用料
						 * 4).設計管制料號by機種入BOM卡關 
						 * *******************************************************************************************/	
						logger.debug("REV = " + affNewRev);
						
						if(rlAdded || rlModified || (initBOMRelease && !rlRemoved)){
							IItem bomItemObject = (IItem)session.getObject(IItem.OBJECT_TYPE, bomItemNo);
							

							/* ===============================================================
							 * 20210813 註解備份: 是否專案用料 與 允許非戰鬥機種使用專案用料的 欄位要取消 
							 * =============================================================== */
							/*
							//不是戰鬥BOM, 也不是車用BOM(正常BOM判斷)
							if(! VALUE_IS_Y.equals(projectItemOfBOM_YN) && ! VALUE_IS_Y.equals(carPlugingOfBOM_YN)){								
								//IItem bomItemObject = (IItem)session.getObject(IItem.OBJECT_TYPE, bomItemNo);
								itemCategory = bomItemObject.getValue( ItemConstants.ATT_TITLE_BLOCK_ITEM_CATEGORY ).toString();
								logger.info(">>>>>>>>>>>> " + bomItemObject.getName() + " - Y/N = " + bomItemObject.getValue(new Integer(PROJECT_ITEM_OF_PART_BASEID) ).toString());
								
								if( VALUE_IS_Y.equals( bomItemObject.getValue(new Integer(PROJECT_ITEM_OF_PART_BASEID) ).toString().toUpperCase() ) ){ // 如果是專案用料
									logger.info("專用料:"+bomItemObject.getName());
									//取得 "允許非戰鬥機種使用專案用料"  欄位的機種來檢查是否可以被加入此正常BOM								
									allowNormalBomUseProjectPartOfModel = bomItemObject.getValue(ASKEYPartsConstants.P2_ALLOW_NORMAL_BOM_USE_PROJECT_ITEM_OF_PART).toString();
									if(! "".equals(allowNormalBomUseProjectPartOfModel)){
										arlAllowNormalBomUseProjectPart = new ArrayList<String>( Arrays.asList( allowNormalBomUseProjectPartOfModel.split(";") ) );
									}
									
									if(! arlAllowNormalBomUseProjectPart.contains(bomModelName)){
										ProjectItemOfPartList.add(bomItemNo);
										existBOMError=true;
										logger.info("專案用料ERROR");
									}									
								}//end if 如果是專案用料
								
								
								// 如果是車載用料(只判斷 EE/ME/成半品)	
								logger.info("\t\taffItem: 如果是車載用料(只判斷 EE/ME/成半品");
								if(itemCategory.equals(ASKEYPartsConstants.PART_CATEGORY_OF_ELECTRONIC) || 
								   itemCategory.equals(ASKEYPartsConstants.PART_CATEGORY_OF_MECHANICS) ||
								   itemCategory.equals(ASKEYPartsConstants.PART_CATEGORY_OF_FG_SA)){
									if( VALUE_IS_Y.equals( bomItemObject.getValue(new Integer(CAR_PLUGING_OF_PART_BASEID) ).toString().toUpperCase() ) ){ 
										CarPlugingOfPartList.add(bomItemNo);
										existBOMError=true;	
										logger.info("車載用料ERROR");
									}	
								}//end if (itemCategory)							
								
							}//end if(! VALUE_IS_Y.equals(projectItemOfBOM_YN) && ! VALUE_IS_Y.equals(carPlugingOfBOM_YN))
							 */
							
							//不是車用BOM(正常BOM判斷)
							if(! VALUE_IS_Y.equals(carPlugingOfBOM_YN)){								
								//IItem bomItemObject = (IItem)session.getObject(IItem.OBJECT_TYPE, bomItemNo);
								itemCategory = bomItemObject.getValue( ItemConstants.ATT_TITLE_BLOCK_ITEM_CATEGORY ).toString();
								logger.debug("itemCategory=\t " + itemCategory);								
								
								// 如果是車載用料(只判斷 EE/ME/成半品)	
								logger.debug("\t\taffItem: 如果是車載用料(只判斷 EE/ME/成半品");
								if(itemCategory.equals(ASKEYPartsConstants.PART_CATEGORY_OF_ELECTRONIC) || 
								   itemCategory.equals(ASKEYPartsConstants.PART_CATEGORY_OF_MECHANICS) ||
								   itemCategory.equals(ASKEYPartsConstants.PART_CATEGORY_OF_FG_SA)){
									if( VALUE_IS_Y.equals( bomItemObject.getValue(new Integer(CAR_PLUGING_OF_PART_BASEID) ).toString().toUpperCase() ) ){ 
										CarPlugingOfPartList.add(bomItemNo);
										existBOMError=true;	
										logger.debug("車載用料ERROR");
									}	
								}//end if (itemCategory)							
								
							}//end if(! VALUE_IS_Y.equals(projectItemOfBOM_YN) && ! VALUE_IS_Y.equals(carPlugingOfBOM_YN))
						
						
							//設計管制料號by機種入BOM卡關 (只判斷 EE) : 936取消設計管制
							/*
							if(itemCategory.equals(ASKEYPartsConstants.PART_CATEGORY_OF_ELECTRONIC)){
								logger.info("\t\taffItem: 設計管制料號by機種入BOM卡關 ");
								if(lifeCycle.equals( ASKEYPartsConstants.LIFECYCLEPHASE_OF_COND_DESIGN ) && rlAdded){
									allowModelUsePartsOfDesignControlStatus = bomItemObject.getValue(ASKEYPartsConstants.P2_ALLOW_MODEL_USE_PARTS_OF_DESIGN_CONTROL_STATUS).toString();
									System.out.println("Item model=" + allowModelUsePartsOfDesignControlStatus);
									System.out.println("BOM Model = " + bomModelName);
									if(! "".equals(allowModelUsePartsOfDesignControlStatus)){
										arlAllowModelUsePartsOfDesignControlStatus = new ArrayList<String>( Arrays.asList( allowModelUsePartsOfDesignControlStatus.split(";") ) );
									}
									if(! arlAllowModelUsePartsOfDesignControlStatus.contains(bomModelName)){
										System.out.println("不包含");
										DesignControlOfPartList.add(bomItemNo);
										existBOMError=true;
										logger.info("設計管制料號 ERROR");
									}		
								}
							}
							*/
							
						}//end if(rlAdded || rlModified || (newRelease && !rlRemoved))
												

						if (!rlRemoved) {
							if (!"".equals(primarySubstitute) ) {
								if(ASKEYBOMConstants.BOM_PRIMARY_SUBSTITUTE_VALUEOF_M.equals(primarySubstitute)){
									//確認是否重複主料
									logger.debug("\t\taffItem: 確認是否重複主料");
									bomItemPrimarySubstitute = bomItemNo + "|" + primarySubstitute+ "|" +boardLayer;
									if (PrimaryList.contains(bomItemPrimarySubstitute)) {
										duplicatePrimaryList.add(bomItemNo);
										existBOMError=true;
									}else{
										PrimaryList.add(bomItemPrimarySubstitute);
									}
								}else{
									//確認是否重複替代料
									logger.debug("\t\taffItem:確認是否重複替代料");
									bomItemFindNoSubstitute= bomItemNo + "|" + findNo+ "|" +boardLayer;
									if (SubstituteList.contains(bomItemFindNoSubstitute)) {
										duplicateSubstituteList.add(findNo);
										existBOMError=true;
									}else{
										SubstituteList.add(bomItemFindNoSubstitute);
									}
								}
							}

							//主料需判斷是否重複製程
							logger.debug("\t\taffItem:主料需判斷是否重複製程");
							if (!"".equals(boardLayer) && ASKEYBOMConstants.BOM_PRIMARY_SUBSTITUTE_VALUEOF_M.equals(primarySubstitute)) {
								bomItemBoardLayer = bomItemNo + "|" + boardLayer;
								if (BoardLayerList.contains(bomItemBoardLayer)) {
									duplicateBoardLayerList.add(bomItemNo);
									existBOMError=true;
								}else{
									BoardLayerList.add(bomItemBoardLayer);
								}
							}

							//判斷Find Number是否重複
							logger.debug("\t\taffItem:判斷Find Number是否重複");
							if (!"".equals(findNo)&& ASKEYBOMConstants.BOM_PRIMARY_SUBSTITUTE_VALUEOF_M.equals(primarySubstitute)) {
								findNoPrimarySubstitute = findNo + "|" + primarySubstitute;
								if (findNoList.contains(findNoPrimarySubstitute)) {
									duplicateFindNumberList.add(findNo);
									existBOMError=true;
								}else{
									findNoList.add(findNoPrimarySubstitute);
								}
							}

							//判斷是否重複插件位置及插件長度是否超出最大限制
							logger.debug("\t\taffItem:判斷是否重複插件位置及插件長度是否超出最大限制");
							if ( ( refDesigs != null && !"".equals(refDesigs) ) && ASKEYBOMConstants.BOM_PRIMARY_SUBSTITUTE_VALUEOF_M.equals(primarySubstitute)){
								logger.debug("\t\refDesigs:1");
								
								if(refDesigs == null){
									logger.debug("\t\refDesigs value string is NULL ");
								}
								else{
									logger.debug("\t\refDesigs value string is NOT NULL ");
								}
								String[] refDesignatorAry=refDesigs.split(",");
								for(String refDes:refDesignatorAry){
									ArrayList<String> rlist=ListRefDesig(refDes);
									
									//判斷是否重複插件位置
									logger.debug("\t\taffItem:判斷是否重複插件位置");
									refDesigQty+=rlist.size();
									Collection interList=CollectionUtils.intersection(rlist,RefDesigListList);
									if(interList.size()>0){
										for(Object r1:interList){
											duplicateRefDesigList.add(r1.toString());
										}
										existBOMError=true;
									}
									RefDesigListList=(ArrayList)CollectionUtils.union(RefDesigListList, rlist);
									
									//判斷插件長度是否超出最大限制
									logger.debug("\t\taffItem:判斷插件長度是否超出最大限制");
									for(String refLengthCheck: rlist){
										boolean bolLengthIsExceedsERPLimit = partsutil.CheckRefDesignLengthIsExceedsERPLimit( refLengthCheck );
										if(bolLengthIsExceedsERPLimit){	
											String refdesErrmsg = String.format("零件[%s]的Ref-Des[%s]長度為[%s]", bomItemNo,refLengthCheck,partsutil.getRefDesign_Length());
											refOverMaxLengthList.add(refdesErrmsg);
											existBOMError=true;
										}
										/*
										if(refLengthCheck.length() > 15){
											refOverMaxLengthList.add(bomItemNo + " | " + refLengthCheck);
											existBOMError=true;
										}
										*/
									}
								}
							}
							else{
								logger.debug(" refDisn is null ");
							}

							//驗證不存在主料
							//將主料與替代料的項次/數量/插件/版階/代用方式/是否為群組料的資料存入HashpMap							
							logger.debug("\t\taffItem:驗證不存在主料");
							if (!"".equals(primarySubstitute)){
								qty = rlRow.getValue(ItemConstants.ATT_BOM_QTY).toString();

								if( ASKEYBOMConstants.BOM_PRIMARY_SUBSTITUTE_VALUEOF_M.equals(primarySubstitute)) {
									findNoMap.put(findNo, new String[]{ bomItemNo, 
																		qty,
																		refDesigs,
																		boardLayer,
																		substituteType,
																		isGroupParts});
								}else{
									if(findNoMap.get(findNo)==null)
										findNoMap.put(findNo, new String[]{"","","","","",""});

									substituteList=findNoSubstituteMap.get(findNo);
									if(substituteList==null){
										substituteList=new ArrayList<String[]>();
									}
									substituteList.add(new String[]{ bomItemNo, 
																	 qty,
																	 refDesigs,
																	 boardLayer,
																	 substituteType,
																	 isGroupParts});
									findNoSubstituteMap.put(findNo, substituteList);
								}
							}
						}
						else{
							logger.debug("findNo:"+findNo);
						}

						if (!rlRemoved) {
							//驗證數量是否為零
							//2024/8/15:調整允許BOM數量可以為0
							logger.debug("\t\taffItem:驗證數量是否為零");
							if(qty==null)
								qty = rlRow.getValue(ItemConstants.ATT_BOM_QTY).toString();
							//if ("".equals(qty) || "0".equals(qty)) {
							if ("".equals(qty)) {
								qtyZeroList.add(bomItemNo);
								existBOMError=true;
							}

							//驗證是否為Preliminary Item
							logger.debug("\t\taffItem:驗證是否為Preliminary Item");
							if("Preliminary".equals(lifeCycle)){
								preliminaryComList.add(bomItemNo);
								existBOMError=true;
							}else if("".equals(lifeCycle)){
								if(!"".equals(bomItemNo.trim())){
									nullItemList.add(bomItemNo);
									existBOMError=true;
								}
							}

							revision = rlRow.getValue(ItemConstants.ATT_BOM_ITEM_REV).toString();
							if("".equals(revision)){
								if(!nullItemList.contains(bomItemNo))
									nullItemList.add(bomItemNo);
								existBOMError=true;
							}

							
							logger.debug("\t\taffItem:驗證Find Number是否為空");
							if ("".equals(findNo)|| "0".equals(findNo) || findNo.startsWith("0")){
								findNoEmpList.add(bomItemNo);
								existBOMError=true;
							}

							//驗證 68BOM 製程不能為空
							logger.debug("\t\taffItem:驗證 68BOM 製程不能為空");
							if (is68BOM && "".equals(boardLayer)){
								boardLayerEmpList.add(bomItemNo);
								existBOMError=true;
							}//驗證68製程是否為Top / Bottom
							else if(is68BOM && 
									!(boardLayer.equals(ASKEYPartsConstants.BOM_LAYER_VALUEOF_TOP) || 
									  boardLayer.equals(ASKEYPartsConstants.BOM_LAYER_VALUEOF_BOTTOM)) ){
								logger.debug("\t\taffItem:驗證68製程是否為Top / Bottom");
								error68BOMLayerList.add(bomItemNo);
								existBOMError=true;							
							}
							
							//驗證 70BOM 製程不為空時,只能填CSP/COB/Boxbuild
							/* 2020/8/25 先註解,因為IE沒有人測試,所以未上線
							if( is70BOM && !"".equals(boardLayer) && 
									!(boardLayer.equals(ASKEYPartsConstants.BOM_LAYER_VALUEOF_COB) ||
									  boardLayer.equals(ASKEYPartsConstants.BOM_LAYER_VALUEOF_CSP) ||
									  boardLayer.equals(ASKEYPartsConstants.BOM_LAYER_VALUEOF_BOX_BUILD))) {
								error70BOMLayerList.add(bomItemNo);
								existBOMError=true;
							}							

							//驗證非68/70 BOM 的製程應該為空
							if (!is68BOM && !is70BOM && !"".equals(boardLayer)){
								boardLayerNotNullList.add(bomItemNo);
								existBOMError=true;
							}
							*/
							
							
							//驗證非68 BOM 的製程應該為空
							logger.debug("\t\taffItem:驗證非68 BOM 的製程應該為空");
							if (!is68BOM && !"".equals(boardLayer)){
								boardLayerNotNullList.add(bomItemNo);
								existBOMError=true;
							}

							//驗證主替代為空的料號
							logger.debug("\t\taffItem:驗證主替代為空的料號");
							if ("".equals(primarySubstitute)){
								primarySubstituteEmptyList.add(bomItemNo);
								existBOMError=true;
							}else if(!ASKEYBOMConstants.BOM_PRIMARY_SUBSTITUTE_VALUEOF_M.equals(primarySubstitute)){
								//替代料驗證代用方式不可為空
								logger.debug("\t\taffItem:替代料驗證代用方式不可為空");
								substituteType = rlRow.getValue(ASKEYPartsConstants.BOM_SUBSTITUTE_TYPE).toString();
								if("".equals(substituteType)){
									substituteTypeEmpList.add(bomItemNo);
									existBOMError=true;
								}
							}else{
								//主料驗證代用方式必須為空
								logger.debug("\t\taffItem:主料驗證代用方式必須為空");
								substituteType = rlRow.getValue(ASKEYPartsConstants.BOM_SUBSTITUTE_TYPE).toString();
								if(!"".equals(substituteType)){
									substituteTypenNotEmpList.add(bomItemNo);
									existBOMError=true;
								}
							}

							//判斷是否插件位置與數量是否一致
							logger.debug("\t\taffItem:判斷是否插件位置與數量是否一致");
							if (refDesigQty>0  && ASKEYBOMConstants.BOM_PRIMARY_SUBSTITUTE_VALUEOF_M.equals(primarySubstitute))
								if (CommonUtil.isInteger(qty)){
									if (refDesigQty!= Integer.parseInt(qty)){
										logger.debug(refDesigQty);
										refQtyErrorList.add(bomItemNo);
										existBOMError=true;
									}
								}
								else{
									refQtyErrorList.add(bomItemNo);
									existBOMError=true;
								}
						    }						
					}//end while					
				}
				catch(Exception ae){
					ExceptionUtil exceptionUtil = new ExceptionUtil();			
					String aeMsg = "StackTrace(ex) : " + exceptionUtil.getExceptionInfo(ae, this.getClass().getName()) + "。" + "Message : " + ae.getMessage() + "。";	
					throw new Exception(aeMsg);
					
				}//Affected Items EndLoop
						
				
				/* ==使用FindNumber處理並檢查 ==
				 * 1.數量
				 * 2.插件
				 * 3.版階
				 * 4.群組料號RULE
				 * ==========================*/				
				Iterator iter = findNoMap.entrySet().iterator(); 
				while (iter.hasNext()) { 
					Map.Entry entry = (Map.Entry) iter.next(); 
					String[] valAry = (String[])entry.getValue(); // qty,ref,boardlayer,substituteType,groupipn
					String key = entry.getKey().toString();       //findnum
					
					if("".equals(valAry[IDX_CHECK_BOM_ITEM])){
						nullPrimaryList.add(key);
						existBOMError=true;
						
					}else{						
						substituteList=findNoSubstituteMap.get(key);
						if(substituteList!=null){
							boolean bolSubstituteIsGroupParts = false;
							
							for(String[] substituteInfo:substituteList){
								if(substituteInfo[IDX_CHECK_BOM_ITEM].equals(valAry[IDX_CHECK_BOM_ITEM])){
									if(!theSamePrimarySubstituteList.contains(key))
										theSamePrimarySubstituteList.add(key);
									existBOMError=true;
								}

								if(!substituteInfo[IDX_CHECK_QTY].equals(valAry[IDX_CHECK_QTY])){
									inconsisentQtyList.add(substituteInfo[0]);
									existBOMError=true;
								}
								
								if(!substituteInfo[IDX_CHECK_REFSESIGNATOR].equals(valAry[IDX_CHECK_REFSESIGNATOR])){	
									inconsisentRefDesigList.add(substituteInfo[0]);
									existBOMError=true;
								}
								
								if(!substituteInfo[IDX_CHECK_BOARD_LAYRER].equals(valAry[IDX_CHECK_BOARD_LAYRER])){	
									inconsisentBoardLayerList.add(substituteInfo[0]);
									existBOMError=true;
								}
								
								//判斷替料為群組料:2023/11/8
								if( substituteInfo[IDX_CHECK_IS_GROUP_PARTS].equals( ASKEYPartsConstants.PART_VALUEOF_Yes ) &&
									(substituteInfo[IDX_CHECK_SUBSTITUTE_TYPE].equals( ASKEYBOMConstants.BOM_SUBSTITUTE_TYPE_VALUEOF_S ))){
									bolSubstituteIsGroupParts = true;
								}
								
							}//end for
							
							//替料若有群組料,主料若為非群組料,則要卡關:2023/11/8
							if(bolSubstituteIsGroupParts && !valAry[IDX_CHECK_IS_GROUP_PARTS].equals( ASKEYPartsConstants.PART_VALUEOF_Yes ) ){
								groupPartsRuleIsViolatedList.add( valAry[IDX_CHECK_BOM_ITEM] + "[項次:" + key + "]"); //主料[find number]
								existBOMError=true;
							}
						}
					}					
				}//end while				

				if(existBOMError){
					if (!"".equals(resultMsg.toString()))
						resultMsg.append("\nBOM[").append(affItem).append("]存在如下錯誤");
					else
						resultMsg.append("BOM[").append(affItem).append("]存在如下錯誤");

					if (primarySubstituteEmptyList.size()>0) {
						resultMsg.append(ERROR_NULL_PRIMARY_SUBSTITUE);
						for(String primarySubstituteStr:primarySubstituteEmptyList){
							resultMsg.append(primarySubstituteStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					//preliminaryComList
					if (preliminaryComList.size()>0) {
						int i=0;
						for(String preliminaryComStr:preliminaryComList){
							if(!preliminaryBOMList.contains(preliminaryComStr)){
								i+=1;
								if(i==1)
									resultMsg.append(ERROR_PRELIMINARY_ITEM);
								resultMsg.append(preliminaryComStr).append(",");
							}
						}
						if(i>0) resultMsg.setLength(resultMsg.length() - 1);
					}

					if(nullPrimaryList.size()>0){
						resultMsg.append(ERROR_NULL_PRIMARY);
						for(String findNoStr:nullPrimaryList){
							resultMsg.append(findNoStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);

					}

					if (findNoEmpList.size()>0) {
						resultMsg.append(ERROR_NULL_FINDNO);
						for(String findNoStr:findNoEmpList){
							resultMsg.append(findNoStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					if (nullItemList.size()>0) {
						resultMsg.append(ERROR_NULL_ITEM);
						for(String itemNo:nullItemList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					if (substituteTypeEmpList.size()>0) {
						resultMsg.append(ERROR_NULL_SUBSTITUTE_TYPE);
						for(String bomItemStr:substituteTypeEmpList){
							resultMsg.append(bomItemStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					//substituteTypenNotEmpList

					if (substituteTypenNotEmpList.size()>0) {
						resultMsg.append(ERROR_NULL_PRIMARY_TYPE);
						for(String bomItemStr:substituteTypenNotEmpList){
							resultMsg.append(bomItemStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}


					if (duplicateFindNumberList.size()>0) {
						resultMsg.append(ERROR_DUPLICATE_FIND_NUMBER);
						for(String duplicateFindNumberStr:duplicateFindNumberList){
							resultMsg.append(duplicateFindNumberStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					if (duplicatePrimaryList.size()>0) {
						resultMsg.append(ERROR_DUPLICATE_PRIMARY);
						for(String duplicatePrimaryStr:duplicatePrimaryList){
							resultMsg.append(duplicatePrimaryStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					if (duplicateSubstituteList.size()>0) {
						resultMsg.append(ERROR_DUPLICATE_SUBSTITUTE);
						for(String duplicateSubstituteStr:duplicateSubstituteList){
							resultMsg.append(duplicateSubstituteStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					if (duplicateBoardLayerList.size()>0) {
						resultMsg.append(ERROR_DUPLICATE_BOARD_LAYRER);
						for(String duplicateBoardLayerStr:duplicateBoardLayerList){
							resultMsg.append(duplicateBoardLayerStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					if (qtyZeroList.size()>0) {
						resultMsg.append(ERROR_NULL_QTY);
						for(String qtyZeroStr:qtyZeroList){
							resultMsg.append(qtyZeroStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					if (refQtyErrorList.size()>0) {
						resultMsg.append(ERROR_REFSESIGNATOR_QTY);
						for(String refQtyErrorStr:refQtyErrorList){
							resultMsg.append(refQtyErrorStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					if (boardLayerEmpList.size()>0) {
						resultMsg.append(ERROR_NULL_BOARD_LAYRER);
						for(String itemNo:boardLayerEmpList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					if (boardLayerNotNullList.size()>0) {
						resultMsg.append(ERROR_NOT_NULL_BOARD_LAYRER);
						for(String itemNo:boardLayerNotNullList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}


					if (inconsisentQtyList.size()>0) {
						resultMsg.append(ERROR_INCOSISENT_QTY);
						for(String itemNo:inconsisentQtyList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					if (theSamePrimarySubstituteList.size()>0) {
						resultMsg.append(ERROR_THE_SAME_PRIMARY_SUBSTITUTE);
						for(String findNumber:theSamePrimarySubstituteList){
							resultMsg.append(findNumber).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					if (inconsisentRefDesigList.size()>0) {
						resultMsg.append(ERROR_INCOSISENT_REFSESIGNATOR);
						for(String itemNo:inconsisentRefDesigList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					if (inconsisentBoardLayerList.size()>0) {
						resultMsg.append(ERROR_INCOSISENT_BOARD_LAYRER);
						for(String itemNo:inconsisentBoardLayerList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}

					//if(showWarning&& duplicateRefDesigList.size()>0){
					if(duplicateRefDesigList.size()>0){
						resultMsg.append(WARNING_DUPLICATE_REFSESIGNATOR);
						for(String duplicateRefDesigStr:duplicateRefDesigList){
							resultMsg.append(duplicateRefDesigStr).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					
					//判斷插件長度是否超出最大限制
					if(refOverMaxLengthList.size()>0){
						resultMsg.append(ERROR_REF_OVER_MAX_LENGTH);
						for(String refLengthChack:refOverMaxLengthList){
							resultMsg.append(refLengthChack).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					
					//顯示專案用料的Warning msg
					/*
					if(ProjectItemOfPartList.size()>0) {
						resultMsg.append(WARNING_PROJECT_ITEM_OF_PART);
						for(String itemNo:ProjectItemOfPartList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}//end 專案用料
					*/
					//顯示車載用料的Warning msg
					if(CarPlugingOfPartList.size()>0){
						resultMsg.append(WARNING_CAR_PLUGING_OF_PART);
						for(String itemNo:CarPlugingOfPartList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}//end  車載用料
					
					//顯示違反群組料規則:2023/11/8
					if(groupPartsRuleIsViolatedList.size() > 0){
						resultMsg.append(ERROR_GROUPIPN_RULE_IS_VIOLATED);
						for(String itemNo:groupPartsRuleIsViolatedList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					
					//顯示設計管制料號不允許入BOM的Warning msg
					/*
					if(DesignControlOfPartList.size() > 0){						
						resultMsg.append(WARNING_DESIGN_CONTROL_OF_PART);
						for(String itemNo:DesignControlOfPartList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}//end 設計管制料號
					*/
					
					/* 未上線
					//顯示68Bom階  Item 製程欄位的Error msg
					if(error68BOMLayerList.size() >0){
						resultMsg.append(ERROR_68BOM_LAYER);
						for(String itemNo:error68BOMLayerList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					
					
					//顯示70Bom階  Item 製程欄位的Error msg
					if(error70BOMLayerList.size() >0){
						resultMsg.append(ERROR_70BOM_LAYER);
						for(String itemNo:error70BOMLayerList){
							resultMsg.append(itemNo).append(",");
						}
						resultMsg.setLength(resultMsg.length() - 1);
					}
					*/
					
				}// end existBOMError
			}
			
			RV_doIt = resultMsg.toString();	
			return RV_doIt;
			
		}
		catch (Exception e) {
			ExceptionUtil exceptionUtil = new ExceptionUtil();			
			String exMsg = "StackTrace(ex) : " + exceptionUtil.getExceptionInfo(e, this.getClass().getName()) + "。" + "Message : " + e.getMessage() + "。";			
			logger.error("doIt Exception:" + exMsg);
			throw new Exception( exMsg );
		}
		finally{			
			resultMsg.setLength(0);//清空String
			//arlAllowNormalBomUseProjectPart.clear();
			arlAllowModelUsePartsOfDesignControlStatus.clear();			
			primarySubstituteEmptyList.clear();
			preliminaryComList.clear();
			preliminaryBOMList.clear();
			refQtyErrorList.clear();
			findNoEmpList.clear();
			nullItemList.clear();
			substituteTypeEmpList.clear();
			substituteTypenNotEmpList.clear();
			boardLayerEmpList.clear();
			boardLayerNotNullList.clear();
			duplicateFindNumberList.clear();
			findNoList.clear();
			qtyZeroList.clear();
			qtyZeroList.clear();
			duplicatePrimaryList.clear();
			duplicateSubstituteList.clear();
			SubstituteList.clear();
			duplicateBoardLayerList.clear();
			BoardLayerList.clear();
			duplicateRefDesigList.clear();
			RefDesigListList.clear();
			findNoMap.clear();
			nullPrimaryList.clear();
			findNoSubstituteMap.clear();
			inconsisentQtyList.clear();
			theSamePrimarySubstituteList.clear();
			inconsisentRefDesigList.clear();
			inconsisentBoardLayerList.clear();
			//ProjectItemOfPartList.clear();
			CarPlugingOfPartList.clear();
			//DesignControlOfPartList.clear();
			error68BOMLayerList.clear();
			error70BOMLayerList.clear();			
					
			logger.debug("Clear Collection Object Success.....");			
		}
	}


//	private static ArrayList<String> ListRefDesig(String refDesignator){
	private ArrayList<String> ListRefDesig(String refDesignator){
		int count=1;
		ArrayList<String> refDesigList=new ArrayList<String>();
		int indexDash = refDesignator.indexOf("~");
		if(indexDash!=-1){
			String a=refDesignator.substring(0,indexDash);
			String b=refDesignator.substring(indexDash+1,refDesignator.length());
			String aPrefix="";
			String bPrefix="";
			String aSubfix="";
			String bSubfix="";
			byte[] aBytes = a.getBytes();
			String aNumber="";
			if(aBytes[aBytes.length-1]<48||aBytes[aBytes.length-1]>57){
				for(byte abyte:aBytes){
					String aString=(char)abyte+"";
					int aAsc2=(int)abyte;
						if(aAsc2<48 || aAsc2>57){
							if(!"".equals(aNumber)){
								//aPrefix=aPrefix+aNumber+aString;
								aSubfix+=aString;
							}else{
								aPrefix+=aString;
							}
							aNumber="";
						}else{
							aNumber+=aString;
						}
					logger.debug("aPrefix="+aPrefix);
					logger.debug("aSubfix="+aSubfix);
					}
					logger.info("aPrefix:"+aPrefix);
					logger.debug("a="+a);
					a=a.replaceAll(aPrefix,"");
					a=a.replaceAll(aSubfix,"");
					logger.debug("a="+a);
				
			}else{
			    for(byte abyte:aBytes){
			    	String aString=(char)abyte+"";
			    	int aAsc2=(int)abyte;
					if(aAsc2<48 || aAsc2>57){
						if(!"".equals(aNumber)){
							aPrefix=aPrefix+aNumber+aString;
							//aSubfix+=aString;
						}else{
							aPrefix+=aString;
						}
						aNumber="";
					}else{
						aNumber+=aString;
					}
					logger.debug("aPrefix="+aPrefix);
					logger.debug("aSubfix="+aSubfix);
				}
				logger.info("aPrefix:"+aPrefix);
				logger.debug("a="+a);
				a=a.replaceAll(aPrefix,"");
				logger.debug("a="+a);
			}
			String bNumber="";
			byte[] bBytes = b.getBytes();
            if(bBytes[bBytes.length-1]<48||bBytes[bBytes.length-1]>57){
            	for(byte bbyte:bBytes){
    				String bString=(char)bbyte+"";

    				int bAsc2=(int)bbyte;
    				if(bAsc2<48 || bAsc2>57){
    					if(!"".equals(bNumber)){
    						//bPrefix=bPrefix+bNumber+bString;
    						bSubfix+=bString;
    					}else{
    						bPrefix+=bString;
    					}
    					bNumber="";
    				}else{
    					bNumber+=bString;
    				}
    			}
            	logger.debug("b="+b);
    			b=b.replaceAll(bPrefix,"");
    			b=b.replaceAll(bSubfix,"");
    			logger.debug("b="+b);
				
			}else{
			for(byte bbyte:bBytes){
				String bString=(char)bbyte+"";

				int bAsc2=(int)bbyte;
				if(bAsc2<48 || bAsc2>57){
					if(!"".equals(bNumber)){
						bPrefix=bPrefix+bNumber+bString;
					}else{
						bPrefix+=bString;
					}
					bNumber="";
				}else{
					bNumber+=bString;
				}
			}
			logger.debug("b="+b);
			b=b.replaceAll(bPrefix,"");
			
			logger.debug("b="+b);
			}
			if(CommonUtil.isInteger(a) && CommonUtil.isInteger(b) && bPrefix.equals(aPrefix)){
				count=Integer.parseInt(b)-Integer.parseInt(a)+1;
			}
			if(count<2){
				refDesigList.add(refDesignator);
			}else{

				for(int i=Integer.parseInt(a);i<Integer.parseInt(b)+1;i++){
					String temp_i=String.valueOf(i);
					logger.debug(aPrefix+addZero(temp_i,a.length())+aSubfix);
					refDesigList.add(aPrefix+addZero(temp_i,a.length())+aSubfix);
				}
			}
		}else{
			logger.debug(refDesignator);
			refDesigList.add(refDesignator);
		}
		return refDesigList;
	}

//	private static int countRefDesig(String refDesignator){
	private int countRefDesig(String refDesignator){
		int count=1;
		int indexDash = refDesignator.indexOf("-");
		if(indexDash!=-1){
			String a=refDesignator.substring(0,indexDash);
			String b=refDesignator.substring(indexDash+1,refDesignator.length());
			String aPrefix="";
			String bPrefix="";
			byte[] aBytes = a.getBytes();
			for(byte abyte:aBytes){
				String aString=(char)abyte+"";
				int aAsc2=(int)abyte;
				if(aAsc2<48 || aAsc2>57){
					a=a.replaceAll(aString,"");
					aPrefix+=aString;
				}
			}
			byte[] bBytes = b.getBytes();
			for(byte bbyte:bBytes){
				String bString=(char)bbyte+"";

				int bAsc2=(int)bbyte;
				if(bAsc2<48 || bAsc2>57){
					b=b.replaceAll(bString,"");
					bPrefix+=bString;
				}
			}
			if(CommonUtil.isInteger(a) && CommonUtil.isInteger(b) && bPrefix.equals(aPrefix)){
				count=Integer.parseInt(b)-Integer.parseInt(a)+1;
			}
			if(count<1)
				count=1; 


		}
		return count;
	}
//	private static  String addZero(String data, int number) {
	private String addZero(String data, int number) {
		String rData = data;
		for (int i = 0; i < number - data.length(); i++) {
			rData = "0" + rData;
		}
		return rData;
	}

	
	//************************************************************************************************
	//***********************************  main entry point    ***************************************
	//************************************************************************************************
	public static void main(String[] args) {
		IAgileSession session = null;
		try{ 	
			
			String keyPath = "c:/temp/secure_key.bin";
        	ASKEYCommonConstants.COMMON_ENCRYPT_GOLDEN_KEY_PATH = keyPath;
			//session = AgileUtil.getAgileSession_TestEnv(); // [PLM 測 試 區]		
			session = AgileUtil.getAgileSession();		   // [PLM 正 試 區] 
			
			String formNumber = "SWAP000001";
			IChange formObject = (IChange) session.getObject(IChange.OBJECT_TYPE, formNumber);
			String error = new PC_15_01_BOMVerification().doIt( session , formObject, true);         
			System.out.println("=error= >>> " + error);
         	System.out.println("====================================");
         	System.out.println("BOMVerifcation Complete !!");         	
        }    
		catch(Exception e){
			ExceptionUtil exceptionUtil = new ExceptionUtil();			
			String errMsg = "StackTrace(ex) : " + exceptionUtil.getExceptionInfo(e, "PC_15_01_BOMVerification") + "。" + "Message : " + e.getMessage() + "。";
			System.out.println("occur exception ....\r\n");
			System.out.println("exception message >>> " + errMsg + " \r\n");
		}		
		finally{
			try{	
				if(session != null && session.isOpen()) session.close();
			}
			catch(APIException ae){			
				ae.printStackTrace();
			}
	    }//End try
	}

}
