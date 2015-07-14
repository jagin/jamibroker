package pl.jagin.jamibroker;

import java.net.URL;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import org.junit.runner.RunWith;

import pl.jagin.jamibroker.junit.Order;
import pl.jagin.jamibroker.JAmiBroker.Stocks;
import pl.jagin.jamibroker.JAmiBroker.Stocks.Stock;
import pl.jagin.jamibroker.junit.OrderedRunner;

import com.jacob.com.ComThread;

import static org.fest.assertions.Assertions.*;

@RunWith(OrderedRunner.class)
// Tests are triggered in order
public class JAmibrokerITest {
	private static String amiBrokerDatabaseDir;
	private static JAmiBroker ab;
	
	@BeforeClass
	public static void beforeClass() throws Exception {
        amiBrokerDatabaseDir = System.getProperty("amiBrokerDatabaseDir", "C:/Program Files (x86)/JAmiBroker/JAmibroker");

		// Initialize the current java thread to be part of the Multi-threaded COM Apartment
		// see: http://stackoverflow.com/questions/980483/jacob-doesnt-release-the-objects-properly
		ComThread.InitMTA();

		ab = JAmiBroker.getInstance();
	}
	
	@AfterClass
	public static void afterClass() {
		ComThread.Release(); // release this java thread from COM
	}
	
	@Test
	@Order(order=10)
	public final void testLoadDatabase() throws Exception {
		boolean result = ab.loadDatabase(amiBrokerDatabaseDir);
		assertThat(result).isEqualTo(true);
	}
	
	@Test
	@Order(order=20)
	public final void testGetDatabasePath() throws Exception {
		assertThat(ab.getDatabasePath()).isEqualTo(amiBrokerDatabaseDir);
	}
	
	@Test
	@Order(order=30)
	public final void testImportFile() throws Exception {
		URL fileUrl = getClass().getResource("/WIG.mst");
		Path filePath = Paths.get(fileUrl.toURI());

		int result = ab.importFile(0, filePath.toAbsolutePath().toString(), "mst.format");
		
		assertThat(result).isEqualTo(0);
		
		Stocks stocks = ab.getStocks();
		Stock stock = stocks.item("WIG");
		
		assertThat(stock).isNotNull();
	}
	
	@Test
	@Order(order=40)
	public final void testStocksAdd() throws Exception {
		Stocks stocks = ab.getStocks();
		
		String ticker = "ABC";
		Stock stock = stocks.add(ticker);
		
		assertThat(stock.getTicker()).isEqualTo(ticker);
	}
	
	@Test
	@Order(order=50)
	public final void testStocksCount() throws Exception {
		Stocks stocks = ab.getStocks();
		
		assertThat(stocks.count()).isEqualTo(2);
	}
	
	@Test
	@Order(order=60)
	public final void testStocksGetTickerList() throws Exception {
		Stocks stocks = ab.getStocks();
		String[] tickers = stocks.getTickerList(0);
		
		assertThat(tickers.length).isEqualTo(2);
	}

	@Test
	@Order(order=70)
	public final void testStocksItem() throws Exception {
		Stocks stocks = ab.getStocks();
		Stock stock = stocks.item(0);
		assertThat(stock.getTicker()).isNotEmpty();
		
		stock = stocks.item("ABC");
		assertThat(stock.getTicker()).isEqualTo("ABC");		
	}
	
	@Test
	@Order(order=80)
	public final void testStocksRemove() throws Exception {
		Stocks stocks = ab.getStocks();
		boolean result = stocks.remove("ABC");

		assertThat(result).isTrue();	
	}
}
