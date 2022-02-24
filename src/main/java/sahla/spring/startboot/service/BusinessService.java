package sahla.spring.startboot.service;

import java.util.Date;
import java.util.HashMap;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import javax.management.OperationsException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;
import org.springframework.web.server.ResponseStatusException;

import sahla.spring.startboot.users.ItemData;
import sahla.spring.startboot.users.RegionData;
import sahla.spring.startboot.users.User;

@Service
public class BusinessService {
	
	SimpleDateFormat format = new SimpleDateFormat("dd-MM-yy");
	static XSSFWorkbook workbook;	

	public List<User> readExcelData () throws OperationsException, ParseException, IOException {
		workbook = new XSSFWorkbook("./data/Users.xlsx");
		XSSFSheet sheet = workbook.getSheetAt(0);
		List<User> users = new ArrayList<User>();

		for(int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
			XSSFRow row = sheet.getRow(i);
			
			if (row.getCell(0) != null && row.getCell(0).toString().trim().length() > 0) {
				User user = new User();
				user.setDate(getCellData(row.getCell(0)));
				user.setRegion(getCellData(row.getCell(1)));
				user.setRep(getCellData(row.getCell(2)));
				user.setItem(getCellData(row.getCell(3)));
				user.setUnits(getCellData(row.getCell(4)));
				user.setUnitCost(getCellData(row.getCell(5)));
				user.setTotal(String.valueOf(row.getCell(6).getNumericCellValue()));
				
				users.add(user);
			}
		}
		
		return users;
	}
	

	public ItemData getItemInfoByDate (String date) throws OperationsException, ParseException, IOException {
		
		ItemData data = null;
		Date inDate = validateDate(date);
		List<User> users = readExcelData();
		
		for (User u : users) {
			String fmtDt = getFormattedDate(u.getDate());
			Date ud = format.parse(fmtDt);
			if (ud.compareTo(inDate) == 0) {
				ItemData itemDet = new ItemData();
				itemDet.setItem(u.getItem());
				itemDet.setUnits(u.getUnits());
				itemDet.setTotal(u.getTotal());
				
				data = itemDet;
			} 
		}
		
		if (data != null)
			return data;
		else
			throw new ResponseStatusException(HttpStatus.NOT_FOUND);
	}
	
	public List<RegionData> getItemsRevenueByRegion () throws OperationsException, ParseException, IOException {
		List<User> users = readExcelData();

		Map<String, User> map = new HashMap<String, User>();
		
		for (User u : users) {
			String key = u.getRegion();
			if (map.get(key) == null) {
				map.put(key, u);
			} else {
				User pu = map.get(key);
				Double tot = Double.valueOf(pu.getTotal()) + Double.valueOf(u.getTotal());
				Double roundedTot = BigDecimal.valueOf(tot).setScale(2, RoundingMode.HALF_UP).doubleValue();
				pu.setTotal(String.valueOf(roundedTot));
			}
		}
		
		List<RegionData> dataList = new ArrayList<RegionData>();

		for (Entry<String, User> entry : map.entrySet()) {
			String key = entry.getKey();
			User user = entry.getValue();
			RegionData regData = new RegionData();
			regData.setRegion(key);
			regData.setTotal(user.getTotal());
			dataList.add(regData);
		}
		return dataList;
	}
		
	private String getFormattedDate(String sd) {
		DecimalFormat df = new DecimalFormat("00");
		String dt = sd;
		String[] vals = sd.split("/");
		if (vals != null && vals.length > 0) {
			String d = vals[0], m = vals[1], y = vals[2];
			if (d.length() < 2) {
				int dv = Integer.valueOf(d);
				d = dv > 9 ? d : df.format(dv);
			}
			if (m.length() < 2) {
				int mv = Integer.valueOf(m);
				m = mv > 9 ? m : df.format(mv);
			}   
			dt = d + "-" + m + "-" + y;
		}
		return dt;
	}

	private Date validateDate(String strDate) throws ParseException {
		Date date = format.parse(strDate);
		return date;
	}

	private String getCellData(XSSFCell cell) {
		DataFormatter formatter = new DataFormatter(); 
		return formatter.formatCellValue(cell);
	}
}

	