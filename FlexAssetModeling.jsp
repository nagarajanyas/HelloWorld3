<%@ taglib prefix="cs" uri="futuretense_cs/ftcs1_0.tld"
%><%@ taglib prefix="ics" uri="futuretense_cs/ics.tld"
%><%@ taglib prefix="render" uri="futuretense_cs/render.tld"
%><%@ taglib prefix="satellite" uri="futuretense_cs/satellite.tld"
%><%@ taglib prefix="asset" uri="futuretense_cs/asset.tld"
%><%@ page import="COM.FutureTense.Interfaces.*,COM.FutureTense.Util.ftMessage,COM.FutureTense.Util.ftErrors,java.io.FileOutputStream,com.fatwire.assetapi.def.AssetTypeDefManager,org.apache.poi.hssf.usermodel.HSSFSheet,org.apache.poi.xssf.usermodel.XSSFSheet,org.apache.poi.ss.usermodel.Sheet,org.apache.poi.hssf.usermodel.HSSFWorkbook,org.apache.poi.xssf.usermodel.XSSFWorkbook,org.apache.poi.ss.usermodel.CreationHelper,org.apache.poi.ss.usermodel.Hyperlink,org.apache.poi.ss.usermodel.CellStyle,org.apache.poi.ss.usermodel.Font,com.fatwire.system.Session,com.fatwire.system.SessionFactory,com.fatwire.assetapi.data.AssetDataManager,java.util.ArrayList,java.util.Map,java.util.HashMap,java.util.Set,java.util.LinkedHashMap,com.fatwire.assetapi.data.AssetId,com.fatwire.assetapi.data.BlobObject,com.fatwire.assetapi.data.AssetData,com.fatwire.assetapi.def.AttributeTypeEnum,java.util.List,com.fatwire.assetapi.query.SimpleQuery,com.fatwire.assetapi.query.Query,java.io.FileInputStream,org.apache.poi.ss.usermodel.Row,org.apache.poi.ss.usermodel.Cell,java.util.Iterator,org.apache.poi.util.IOUtils,java.io.InputStream,java.io.OutputStream,org.apache.poi.ss.usermodel.Drawing,org.apache.poi.ss.usermodel.ClientAnchor,org.apache.poi.ss.usermodel.Picture,java.io.File,org.apache.poi.ss.usermodel.Workbook,org.apache.poi.ss.usermodel.WorkbookFactory,com.fatwire.assetapi.def.AssetTypeDef,com.fatwire.assetapi.def.AttributeDef,com.fatwire.assetapi.def.AssetAssociationDef,com.openmarket.xcelerate.asset.AssetIdImpl,com.fatwire.assetapi.data.AttributeData,org.apache.log4j.Logger,java.net.URL,java.io.IOException,javax.servlet.jsp.JspWriter,au.com.bytecode.opencsv.CSVWriter,java.io.FileWriter,java.net.URLEncoder,java.nio.channels.FileChannel"%>
<cs:ftcs>
<%-- 
/ExportFromOWCS
INPUT
OUTPUT
12345
--%>
<%-- Record dependencies for the Template --%>
<ics:if condition='<%=ics.GetVar("tid")!=null%>'><ics:then><render:logdep cid='<%=ics.GetVar("tid")%>' c="Template"/></ics:then></ics:if>
<html>
	<render:calltemplate tname="/Shay_Head" args="c,cid" >		
	</render:calltemplate>
	<body>
		<div id="myModal" class="modal">
			<!-- Modal content -->
			<div class="modal-content">
				<span class="close" onclick="windowClose();">&times;</span>
				<p id="subType"></p>
				<p id="userTypeId"></p>
				<input type="text" name="name" id="nameTxt">
				<input type="button" value="ok" id="textDisplay" onclick="windowClose();">				
			</div>
		</div><%
		Logger log = Logger.getLogger("Entering into Export Flex Asset Template");
		try{
			out.println("Just edited");
			// Reading values from property (Config Files) and assigning to local variables
			String defaultFileFormat = "excel";
			String outputFormat = Utilities.goodString(ics.GetProperty("OutputFileFormat","JPMC_Constants.ini",true))? ics.GetProperty("OutputFileFormat","JPMC_Constants.ini",true) : defaultFileFormat ;
			String exportFileLocation = ics.GetProperty("ExcelFileLocation","JPMC_Constants.ini",true);	
			String blobLocation = ics.GetProperty("BlobLocation","JPMC_Constants.ini",true);
			String assetTypes = ics.GetProperty("assetType","JPMC_Constants.ini",true);
			String subTypes = ics.GetProperty("assetSubType","JPMC_Constants.ini",true); 
			String nameFieldSeprator = ics.GetProperty("NameFieldSeprator","JPMC_Constants.ini",true); 
			//Creating Session Object and supporting OWCS ADM object
			Session ses = SessionFactory.getSession();
			AssetDataManager assetDataMgr = (AssetDataManager)ses.getManager(AssetDataManager.class.getName());
			AssetTypeDefManager defManager = (AssetTypeDefManager)ses.getManager(AssetTypeDefManager.class.getName());
			
			List<String> assetTypeList = getAssetTypeList(assetTypes);
			if(assetTypeList == null ){%>
				<script>
					ErrorMessage("Please Provide the Appropriate AssetType in INI File");
				</script><%
			}
			Map<String,List> assetTypeMap = getAssetTypeMap(assetTypeList,subTypes,ics);
			Set<String> assetTypeSet = assetTypeMap.keySet();
			for(String assetType:assetTypeSet){
				List<String> assetSubTypeList = assetTypeMap.get(assetType);
				Map<String,List> valueMap = new LinkedHashMap<String,List>();	
				Map<String,List> headerMap = new LinkedHashMap<String,List>();	
				for(String subType:assetSubTypeList){
					List<String> attributeList = new ArrayList<String>();
					attributeList.add("name");
					attributeList.add("description");
					List<String> associatedAttrList = new ArrayList<String>();
					List<String> associatedParentList = new ArrayList<String>();
					
					AssetTypeDef assetTypeDef = defManager.findByName(assetType,subType);
					List<AttributeDef> attrDef = assetTypeDef.getAttributeDefs();
					List<AssetAssociationDef> associationDefList = defManager.findByName(assetType,subType).getAssociations();
					for(AttributeDef def:attrDef){
						if(!def.isMetaDataAttribute()){
							attributeList.add(def.getName());
						}
					}
					if(!associationDefList.isEmpty()){
						for(AssetAssociationDef assocDef:associationDefList){
							associatedAttrList.add(assocDef.getName());
						}
					}
					List<String> headerList = getHeaderList(attributeList,associatedAttrList,associatedParentList);
					headerMap.put(subType,headerList);
					Query assetQry = new SimpleQuery(assetType,subType,null,attributeList);
					Iterable<AssetData> assetIterator = null;
					assetIterator = assetDataMgr.read(assetQry);
					//List<String[]> contArrayOfList = new ArrayList<String[]>();
					List<List> containerList = new ArrayList<List>();
					List<String[]> containerArrayOfList = new ArrayList<String[]>();
					for(AssetData data:assetIterator){						
						List<String> assetValue = getContentData(assetType,subType,data,attributeList,associatedAttrList,associatedParentList,blobLocation,nameFieldSeprator,ics);
						String assetValArray[] = (String []) assetValue.toArray(new String[assetValue.size()]);
						containerList.add(assetValue);
						containerArrayOfList.add(assetValArray);
					}
					valueMap.put(subType,containerArrayOfList);				
				}
				if("csv".equals(outputFormat)){
					String csvFileLocation = ics.GetProperty("CSVFileLocation","JPMC_Constants.ini",true);	
					//out.println("HeaderMap"+headerMap+"::ValueMap::"+valueMap+":::"+csvFileLocation+"::");
					generateCSV(headerMap,valueMap,csvFileLocation,ics);
					out.println("<h1 style='color:Blue;'>"+" CSV Files are genrated successfully on following Location</h1><BR/>"+ "<p style='color:Blue;'>"+csvFileLocation+"<p>" );
				}else if("excel".equals(outputFormat)){
					//Creating Dummy excel to read type of version excel software installed	
					File attrFile = new File(exportFileLocation+"JPMC_Dummy.xlsx");
					FileInputStream attrInputStream = new FileInputStream(attrFile);
					Workbook workBookObj = createWorkBook(attrInputStream,attrFile);
					//Export the excel to the appropriate location as per config file	
					FileOutputStream fout = new FileOutputStream(exportFileLocation+"JPMC_FlexAssetModelling.xlsx");
					generateExcel(workBookObj,headerMap,valueMap,fout,ics);
					out.println("<h1 style='color:Blue;'>"+" Excel Files are genrated successfully on following Location</h1><BR/>"+ "<p style='color:Blue;'>"+exportFileLocation+"<p>" );
					fout.flush();
					fout.close();
				}
				tabularDisplay(headerMap,valueMap,out);				
			}
		}catch(Exception ex){
			ics.LogMsg("Exception In Tool Execution................................."+ex.getMessage());
		} %>		
	</body>
</html>
</cs:ftcs>
<%!
	/*public static void saveImage(String imageUrl, String destinationFile,ICS ics) throws IOException {		
		try{
			String imageURL  = "";
			imageURL = URLEncoder.encode(imageUrl, "UTF-8");
			URL url = new URL(imageURL);
			InputStream is = url.openStream();
			OutputStream os = new FileOutputStream(destinationFile);
			byte[] b = new byte[2048];
			int length;
			while ((length = is.read(b)) != -1) {
				os.write(b, 0, length);
			}
			is.close();
			os.close();
		}catch(Exception ex){
			//ics.LogMsg("Exception in saveImage method*******************"+ex.getMessage());
		}
	}*/

	public static void copyImage(String sourceName, String destName)throws IOException {
		FileChannel source = null;
		FileChannel destination = null;
		try {
			File sourceFile = new File(sourceName);
			File destFile = new File(destName);
			if (!destFile.exists()) {
				destFile.createNewFile();
			}
			        
            source = new FileInputStream(sourceFile).getChannel();
            destination = new FileOutputStream(destFile).getChannel();
            // previous code: destination.transferFrom(source, 0, source.size());
            // to avoid infinite loops, should be:
            long count = 0;
            long size = source.size();
            while ((count += destination.transferFrom(source, count, size - count)) < size);
        } catch(Exception ex){
			//ics.LogMsg("Exception in copyImage method*******************"+ex.getMessage());
		}finally {
            if (source != null) {
                source.close();
            }
            if (destination != null) {
                destination.close();
            }
        }
    }

	public List getAssetTypeList(String assetTypeStr){
		List<String> assetTypeList = null;
		try{			
			if(Utilities.goodString(assetTypeStr)){
				//If not read info from ini file and covert to list
				assetTypeList = new ArrayList<String>();
				for(String type:assetTypeStr.split(",")){
					assetTypeList.add(type);
				}		
			}			
		}catch(Exception ex){
			//ics.LogMsg("Exception in getAssetTypeList method*******************"+ex.getMessage());
		}
		return assetTypeList;
	}

	public Map getAssetTypeMap(List<String> assetTypeList,String subTypes,ICS ics){
		Map<String,List<String>> assetTypeMap = null;
		try{
			assetTypeMap = new HashMap<String,List<String>>();
			for(String assetType:assetTypeList){
				List<String> subTypeList = new ArrayList<String>();		
				if(IsComplexFamily(ics)){
					if(Utilities.goodString(subTypes)){
						subTypeList = getSubTypeList(subTypes);		
					}
				}else{
					if(Utilities.goodString(subTypes)){
						subTypeList = getSubTypeList(subTypes);		
					}else{
						try{
							Session ses = SessionFactory.getSession();
							AssetTypeDefManager defManager = (AssetTypeDefManager)ses.getManager(AssetTypeDefManager.class.getName());
							subTypeList = defManager.getSubTypes(assetType);
						}catch(Exception ex){
							//ics.LogMsg("Exception in getAssetTypeMap method while creating subtypelist object::: "+ex.getMessage());
						}
					}
				}
				assetTypeMap.put(assetType,subTypeList);
			}			
		}catch(Exception ex){
			//ics.LogMsg("Exception in getAssetTypeMap method*******************"+ex.getMessage());
		}
		return assetTypeMap;
	}
	
	public boolean IsComplexFamily(ICS ics){
		boolean complexFamily = false;
		try{
			complexFamily = Boolean.parseBoolean(ics.GetProperty("ComplexFamily","JPMC_Constants.ini",true));			
		}catch(Exception ex){
			//ics.LogMsg("Exception in IsComplexFamily method*******************"+ex.getMessage());
		}
		return complexFamily;			
	}

	public List getSubTypeList(String subTypes){
		List<String> subTypeList = null;
		try {
			subTypeList = new ArrayList<String>();		
			for(String type:subTypes.split(",")){
				subTypeList.add(type);
			}			
		}catch(Exception ex){
			//ics.LogMsg("Exception in getSubTypeList method*******************"+ex.getMessage());
		}
		return subTypeList;
	}
	public List<String> getHeaderList(List<String> attributeList,List<String> associatedAttrList,List<String> parentList){
		List<String> headerList = null;
		try{
			headerList = new ArrayList(); 
			headerList.add("AssetID");
			headerList.add("AssetType");
			headerList.add("Subtype");
			for(String attr:attributeList){
				headerList.add(attr);
			}
			for(String assoAttr:associatedAttrList){
				headerList.add(assoAttr);
			}
			for(String parent:parentList){
				headerList.add("Parents");
			}			
		}catch(Exception ex){
			//ics.LogMsg("Exception in getHeaderList method*******************"+ex.getMessage());
		}
		return headerList;
	}
	public String getAssetName(String assetType, String assetId)throws Exception {
		String assetName = "";
		try{
			Session ses = SessionFactory.getSession();
			AssetDataManager mgr =(AssetDataManager) ses.getManager( AssetDataManager.class.getName() );
			AssetId id = new AssetIdImpl( assetType, Long.parseLong(assetId) );
			List attrNames = new ArrayList();
			attrNames.add( "name" );
			AssetData data = mgr.readAttributes( id, attrNames );
			assetName = data.getAttributeData( "name" ).getData().toString();			
		}catch(Exception ex){
			//ics.LogMsg("Exception in getAssetName method*******************"+ex.getMessage());
		}
		return assetName;
	}
	
	public List<String> getContentData(String assetType,String subType,AssetData data,List<String> attributeList,List<String> associatedAttrList,List<String> parentList,String blobLocation,String nameFieldSeprator,ICS ics) throws Exception{
		List<String> valueList = null;
		try{
			String assetId = Long.toString(data.getAssetId().getId());
			valueList = new ArrayList();
			valueList.add(assetId);
			valueList.add(assetType);
			valueList.add(subType);
			for(String attr:attributeList){			
				String assetVal = "";
				//ics.LogMsg(attr+":::"+data.getAttributeData(attr).getType()+"***</br>");
				if(data.getAttributeData(attr).getType().equals(AttributeTypeEnum.BLOB)){ 
					String fileName = "";
					String folderName = "";
					AttributeData blobAttribute = data.getAttributeData(attr);
					BlobObject fileObj = (BlobObject)blobAttribute.getData();
					if(fileObj != null){
						fileName = fileObj.getFilename();
						folderName = fileObj.getFoldername().replace("/","\\");
						String imageURL = folderName + fileName;
						ics.LogMsg("ImageURL*********"+imageURL);					
						String destinationFile = blobLocation+getAssetName(assetType,assetId)+".jpg";
						//saveImage(imageURL, destinationFile,ics);
						copyImage(imageURL,destinationFile);									
						valueList.add(destinationFile);
					}else{ 
						ics.LogMsg("Associate Any Blob Value for the attribute:::"+attr+"::in the Asset Of::"+assetId);
					}
				}else if(data.getAttributeData(attr).getType().equals(AttributeTypeEnum.ASSET)){
					StringBuffer sb = new StringBuffer();
					List<AssetId> assetList = new ArrayList<AssetId>();
					assetList = data.getAttributeData(attr).getDataAsList();
					for(AssetId eachAsset:assetList){
						String eachAssetType = eachAsset.getType();
						String eachAssetId = Long.toString(eachAsset.getId()); 
						//load and get asset name using asset api
						String assetName = getAssetName(eachAssetType,eachAssetId);
						sb.append(assetName+"|"+eachAssetId+";");
					}
					if(assetList.size() > 0){
						sb.deleteCharAt(sb.length()-1);
					}
					valueList.add(sb.toString());				
				}else{
					String tempCellVal = data.getAttributeData(attr).getDataAsList() != null ? data.getAttributeData(attr).getDataAsList().toString() : "" ;
					tempCellVal=tempCellVal.replaceAll(","," ");
					if(Utilities.goodString(tempCellVal)){
							tempCellVal = tempCellVal.substring(1,tempCellVal.length()-1);								
					}
					if(attr.equalsIgnoreCase("name")){
						//To Display Name Field With AssetId For avoid Duplicate in Drupal while Importing Asset
						String nameFieldWithId = tempCellVal+nameFieldSeprator+String.valueOf(data.getAssetId().getId());
						valueList.add(nameFieldWithId);
					}else{															
						valueList.add(tempCellVal);				
					}
				}
			}
			
			for(String assoAttr:associatedAttrList){ 		
				StringBuffer sb = new StringBuffer();
				List<AssetId> assocAssetList = data.getAssociatedAssets(assoAttr);
				for(AssetId assocAsset:assocAssetList){
					String assocType = assocAsset.getType();
					String assocId = Long.toString(assocAsset.getId());
					String assetName = getAssetName(assocType,assocId);
					sb.append(assetName +"|"+assocId+"!");
				}
				if(assocAssetList.size() > 0){
					sb.deleteCharAt(sb.length()-1);
				}
				valueList.add(sb.toString()); 		
			}
			
			StringBuffer parentSB = new StringBuffer();	
			List<AssetId> parentAssetList = data.getImmediateParents();
			for(AssetId parentAsset:parentAssetList){
				String parentType = parentAsset.getType();
				String parentId = Long.toString(parentAsset.getId());
				String assetName = getAssetName(parentType,parentId);
				parentSB.append(assetName +"|"+parentId+"^");
			}
			if(parentAssetList.size() > 0){
				parentSB.deleteCharAt(parentSB.length()-1);
			}
			valueList.add(parentSB.toString());	
		}catch(Exception ex){
			//ics.LogMsg("Exception in getContentData method*******************"+ex.getMessage());
		}
		return valueList;
	}

	public void generateCSV(Map headerMap,Map valueMap,String csvFileLocation,ICS ics)throws IOException{
		Set<String> valueSet = null;
		try{
			valueSet = valueMap.keySet();
			for(String subType:valueSet){
					File outputFile = new File(csvFileLocation+subType+".csv");
					CSVWriter writer = new CSVWriter(new FileWriter(outputFile));
					List<String> subTypeHeadList = (List<String>)headerMap.get(subType);
					List<String []> subTypeConArrOfList = (List<String[]>)valueMap.get(subType);
					String[] subTypeHeaderArray = (String []) subTypeHeadList.toArray(new String[subTypeHeadList.size()]);
					writer.writeNext(subTypeHeaderArray);
					writer.writeAll(subTypeConArrOfList);
					writer.close();					
			}
		}catch(Exception ex){
			//ics.LogMsg("Exception in generateCSV method*******************"+ex.getMessage());
		}
	}
	
	public String getTruncSubType(String subType){
		try{
			if(subType.length() > 31){
				subType = subType.substring(0,31);
			}			
		}catch(Exception ex){
			//ics.LogMsg("Exception in getTruncSubType method*******************"+ex.getMessage());
		}
		return subType;
	}

	public Workbook createWorkBook(FileInputStream attrInputStream,File attrFile)throws Exception {
		Workbook workBookObj = null;
		try{	
			workBookObj = WorkbookFactory.create(attrInputStream);
			Workbook attrFileWorkBook = null;
			if (workBookObj.getSheetAt(0) instanceof HSSFSheet) {
				workBookObj = new HSSFWorkbook();
				attrInputStream = new FileInputStream(attrFile);
				attrFileWorkBook = new HSSFWorkbook(attrInputStream);
			}else if(workBookObj.getSheetAt(0) instanceof XSSFSheet){
				workBookObj = new XSSFWorkbook();
				attrInputStream = new FileInputStream(attrFile);
				attrFileWorkBook = new XSSFWorkbook(attrInputStream);
			}
			/* Style of WorkBook */
			CellStyle hlinkStyle = workBookObj.createCellStyle();
			Font hlinkFont = workBookObj.createFont();
			hlinkFont.setUnderline(Font.U_SINGLE);
			hlinkFont.setColor(Font.BOLDWEIGHT_BOLD);
			hlinkStyle.setFont(hlinkFont); 
			CreationHelper helper = workBookObj.getCreationHelper(); 
			/* End Here */
		}catch(Exception ex){
			//ics.LogMsg("Exception in createWorkBook method*******************"+ex.getMessage());
		}
		return workBookObj;
	}

	public void generateExcel(Workbook workBookObj,Map headerMap,Map valueMap,FileOutputStream fout,ICS ics) throws IOException{
		try{
			int rowCounter = 1;
			int cellCounter = 0;
			Set<String> subTypeSet = headerMap.keySet();
			for(String subType:subTypeSet){
				String subTypeName = getTruncSubType(subType);
				Sheet sheet =  workBookObj.createSheet(subTypeName);
				Row headerRow = sheet.createRow(0);			
				
				CellStyle boldStyle = workBookObj.createCellStyle();
				Font boldFont = workBookObj.createFont();
				boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
				boldStyle.setFont(boldFont);


				List<String> headerAttrList = (List<String>)headerMap.get(subType);
				List<String[]> assetvalueList = (List<String[]>)valueMap.get(subType);
				for(int i = 0; i < headerAttrList.size(); i++ ){
					sheet.setColumnWidth(i, 8000);	
					Cell headerCell = headerRow.createCell((int) cellCounter++);
					headerCell.setCellValue(headerAttrList.get(i));
					headerCell.setCellStyle(boldStyle);
				}
				cellCounter = 0;
				ics.LogMsg("AssetValue List Size:::"+Integer.toString(assetvalueList.size())+"Of AssetSub Type"+subType);
						
				for(String[] tempRow : assetvalueList){
						Row row = sheet.createRow(rowCounter++); 
						for(int i=0;i<tempRow.length;i++){
							Cell cell = row.createCell(cellCounter++); 
							cell.setCellValue(tempRow[i].toString() );
							//ics.LogMsg("SubType***"+subType+"***tempRow*********"+tempRow[i].toString() );						
						}
						cellCounter = 0;
				}
				rowCounter = 1;
				
				cellCounter = 0;	
			}	
			workBookObj.write(fout);
		}catch(Exception ex){
			//ics.LogMsg("Exception in generateExcel method*******************"+ex.getMessage());
		}		
	}

	public void tabularDisplay(Map headerMap,Map valueMap,JspWriter out) throws IOException{
		try{
			Set<String> subTypeSet = headerMap.keySet();
			for(String subType:subTypeSet){
				out.println("<h1 style='text-align:left'>"+subType+"</h1></br>");
				out.println("<table border='1'>");
				//Displaying Header Attributes of an tables
				List<String> headerAttrList = (List<String>)headerMap.get(subType);
				List<String[]> assetvalueList = (List<String[]>)valueMap.get(subType);
				
				out.println("<tr>");
				for(String headerAttribute : headerAttrList){
					out.println("<th>"+headerAttribute+"</th>");
				}
				out.println("</tr>");
				
				for(String[] tempRow : assetvalueList){
					out.println("<tr>");
					for(String rowValue : tempRow){
						out.println("<td>"+rowValue+"</td>");											
					}
					out.println("</tr>");
				}
				out.println("</table>");
			}
		}catch(Exception ex){
			//ics.LogMsg("Exception in tabularDisplay method*******************"+ex.getMessage());
		}
	}
%>
