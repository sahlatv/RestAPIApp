package sahla.spring.startboot.users;

import java.io.IOException;
import java.text.ParseException;
import java.util.List;

import javax.management.OperationsException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.server.ResponseStatusException;

import sahla.spring.startboot.service.BusinessService;

@RestController
public class ExcelController {
	@Autowired
	BusinessService service;
	
	@RequestMapping("/api/v1/item/info/by/date/{date}")
	public ItemData getItemInfoByDate(@PathVariable String date) throws OperationsException {
		ItemData user = null;
		try {
			user = service.getItemInfoByDate(date);
			return user;
		} catch (IOException e) {
			e.printStackTrace();
			throw new ResponseStatusException(HttpStatus.BAD_REQUEST, e.getMessage());
		} catch (ParseException e) {
			e.printStackTrace();
			throw new OperationsException("Invalid : " + e.getMessage());
		} catch (Exception e) {
			e.printStackTrace();
			throw new OperationsException(e.getMessage());
		}
	}
	
	@RequestMapping("/api/v1/items/revenue/by/region")
	public List<RegionData> getItemsRevenueByRegion() throws OperationsException {
		List<RegionData> user = null;
		try {
			XSSFWorkbook workbook = new XSSFWorkbook("./data/Users.xlsx");
			user = service.getItemsRevenueByRegion();
			return user;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			throw new ResponseStatusException(HttpStatus.BAD_REQUEST, e.getMessage());
		} catch (ParseException e) {
			e.printStackTrace();
			throw new OperationsException("Invalid : " + e.getMessage());
		} catch (Exception e) {
			e.printStackTrace();
			throw new OperationsException(e.getMessage());
		}
		
		
	}
}