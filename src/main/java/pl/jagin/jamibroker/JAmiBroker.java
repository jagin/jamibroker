package pl.jagin.jamibroker;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Date;

/**
 * This class is a wrapper for OLE automation interface of AmiBroker application.
 * For more details @see https://www.amibroker.com/guide/objects.html
 */
public class JAmiBroker {
	private static final Logger log = LoggerFactory.getLogger(JAmiBroker.class);
	
	private static JAmiBroker instance;
	
	private ActiveXComponent dAmiBroker;
	
	private JAmiBroker() { dAmiBroker = new ActiveXComponent("Broker.Application"); }
	
	/**
	 * The getInstance() method returns a reference to a JAmiBroker object, which can be used to execute AmiBroker OLE methods.
	 * This method returns a singleton, so calling it twice in a row will return the same instance. 
	 * @return JAmiBroker
	 */
	public static JAmiBroker getInstance() {
		if(instance == null) {
			instance = new JAmiBroker();
		}
		return instance;
	}
	
	public long importFile(int type, String importFileName) {
		Variant vResult = Dispatch.call(dAmiBroker, "Import", 
				new Variant(type),
				new Variant(importFileName));
		return vResult.getLong();
	}
	
	public int importFile(int type, String importFileName, String formatFileName) {
		Variant vResult = Dispatch.call(dAmiBroker, "Import", 
				new Variant(type),
				new Variant(importFileName),
				new Variant(formatFileName));
		return vResult.getInt();
	}
	
	public boolean loadDatabase(String databasePath) {
		Variant vResult = Dispatch.call(dAmiBroker, "LoadDatabase",
				new Variant(databasePath));
		return vResult.getBoolean();
	}
	
	public boolean loadLayout(String layoutFileName) {
		Variant vResult = Dispatch.call(dAmiBroker, "LoadLayout", 
				new Variant(layoutFileName));
		return vResult.getBoolean();
	}	
	
	public void quit() {
		Dispatch.callSub(dAmiBroker, "Quit");
	}	
	
	public void refreshAll() {
		Dispatch.callSub(dAmiBroker, "RefreshAll");
	}
	
	public void saveDatabase() {
		Dispatch.callSub(dAmiBroker, "SaveDatabase");
	}
	
	public boolean saveLayout(String layoutFileName) {
		Variant vResult = Dispatch.call(dAmiBroker, "SaveLayout", 
				new Variant(layoutFileName));
		return vResult.getBoolean();
	}	
	
	public long log(int action) {
		Variant vResult = Dispatch.call(dAmiBroker, "Log", 
				new Variant(action));
		return vResult.getLong();
	}		
	
	public Stocks getStocks() {
		Dispatch dStocks  = dAmiBroker.getProperty("Stocks").toDispatch();
		return new Stocks(dStocks);
	}
	
	public Stocks getAnalysisDocs() {
		Dispatch dAnalysisDocs  = dAmiBroker.getProperty("AnalysisDocs").toDispatch();
		return new Stocks(dAnalysisDocs);
	}
	
	public String getVersion() {
		Variant vResult = dAmiBroker.getProperty("Version");
		return vResult.getString();
	}
	
	public String getDatabasePath() {
		Variant vResult = dAmiBroker.getProperty("DatabasePath");
		return vResult.getString();
	}
	
	public int getVisible() {
		Variant vResult = dAmiBroker.getProperty("Visible");
		return vResult.getInt();
	}
	
	public void setVisible(int visible) {
		dAmiBroker.setProperty("Visible", visible);
	}	

	public class Stocks {
		private Dispatch dStocks;
		
		private Stocks(Dispatch dStocks) {
			this.dStocks = dStocks;
		}
		
		public Stock add(String ticker) {
			Variant vStock = Dispatch.call(dStocks, "Add", new Variant(ticker));
			
			return !vStock.isNull() ? new Stock(vStock.toDispatch()) : null;
		}
		
		public Stock item(String ticker) {
			Variant vStock =  Dispatch.call(dStocks, "Item", new Variant(ticker));
			
			return !vStock.isNull() ? new Stock(vStock.toDispatch()) : null;
		}	
		
		public Stock item(int index) {
			Variant vStock =  Dispatch.call(dStocks, "Item", new Variant(index));
			
			return !vStock.isNull() ? new Stock(vStock.toDispatch()) : null;
		}		
		
		public String[] getTickerList(long type) {
			Variant vTickerList = Dispatch.call(dStocks, "GetTickerList",
					new Variant(type));
			return vTickerList.getString().split(",");
		}
		
		public boolean remove(String ticker) {
			Variant vResult = Dispatch.call(dStocks, "Remove",
					new Variant(ticker));
			return vResult.getBoolean();
		}	
		
		public boolean remove(int index) {
			Variant vResult = Dispatch.call(dStocks, "Remove",
					new Variant(index));
			return vResult.getBoolean();
		}			
		
		public int count() {
			Variant vResult = Dispatch.get(dStocks, "Count");
			
			return vResult.getInt();
		}
		
		public class Stock {
			private Dispatch dStock;
			
			private Stock(Dispatch dStock) {
				this.dStock = dStock;
			}
			
			private Variant get(String propertyName) {
				return Dispatch.get(dStock, propertyName);
			}
			
			private void set(String propertyName, Variant propertyValue) {
				Dispatch.put(dStock, propertyName, propertyValue);
			}
			
			public String getTicker() {
				return get("Ticker").getString();
			}
			
			public void setTicker(String ticker) {
				set("Ticker", new Variant(ticker));
			}
			
			public Quotations getQuotations() {
				Dispatch dQuotations = Dispatch.get(dStock, "Quotations").toDispatch();
				return new Quotations(dQuotations);
			}
			
			public String getFullName() {
				return get("FullName").getString();
			}
			
			public void setFullName(String fullName) {
				set("FullName", new Variant(fullName));
			}
			
			public boolean isIndex() {
				return get("Index").getBoolean();
			}
			
			public void setIndex(boolean index) {
				set("Index", new Variant(index));
			}
			
			public boolean isFavorite() {
				return get("Favorite").getBoolean();
			}
			
			public void setFavorite(boolean favorite) {
				set("Favorite", new Variant(favorite));
			}
			
			public boolean isContinuous() {
				return get("Continuous").getBoolean();
			}
			
			public void setContinuous(boolean continuous) {
				set("Continuous", new Variant(continuous));
			}
			
			public long getMarketID() {
				return get("MarketID").getLong();
			}
			
			public void setMarketID(long marketID) {
				set("MarketID", new Variant(marketID));
			}
			
			public long getGroupID() {
				return get("GroupID").getLong();
			}
			
			public void setGroupID(long groupID) {
				set("GroupID", new Variant(groupID));
			}
			
			public String getBeta() {
				return get("Beta").getString();
			}
			
			public void setBeta(String beta) {
				set("Beta", new Variant(beta));
			}
			
			public double getSharesOut() {
				return get("SharesOut").getDouble();
			}
			
			public void setSharesOut(double sharesOut) {
				set("SharesOut", new Variant(sharesOut));
			}			
			
			public double getBookValuePerShare() {
				return get("BookValuePerShare").getDouble();
			}
			
			public void setBookValuePerShare(double bookValuePerShare) {
				set("BookValuePerShare", new Variant(bookValuePerShare));
			}	
			
			public double getSharesFloat() {
				return get("SharesFloat").getDouble();
			}
			
			public void setSharesFloat(double sharesFloat) {
				set("SharesFloat", new Variant(sharesFloat));
			}
			
			public String getAddress() {
				return get("Address").getString();
			}
			
			public void setAddress(String address) {
				set("Address", new Variant(address));
			}
			
			public String getWebID() {
				return get("WebID").getString();
			}
			
			public void setWebID(String webID) {
				set("WebID", new Variant(webID));
			}
			
			public String getAlias() {
				return get("Alias").getString();
			}
			
			public void setAlias(String alias) {
				set("Alias", new Variant(alias));
			}		
			
			public boolean isDirty() {
				return get("Dirty").getBoolean();
			}
			
			public void setDirty(boolean dirty) {
				set("Dirty", new Variant(dirty));
			}
			
			public long getIndustryID() {
				return get("IndustryID").getLong();
			}
			
			public void setIndustryID(long industryID) {
				set("IndustryID", new Variant(industryID));
			}
			
			public long getWatchListBits() {
				return get("WatchListBits").getLong();
			}
			
			public void setWatchListBits(long watchListBits) {
				set("WatchListBits", new Variant(watchListBits));
			}
			
			public double getWatchListBits2() {
				return get("WatchListBits2").getDouble();
			}
			
			public void setWatchListBits2(double watchListBits2) {
				set("WatchListBits2", new Variant(watchListBits2));
			}			
			
			public long getDataSource() {
				return get("DataSource").getLong();
			}
			
			public void setDataSource(long dataSource) {
				set("DataSource", new Variant(dataSource));
			}
			
			public long getDataLocalMode() {
				return get("DataLocalMode").getLong();
			}
			
			public void setDataLocalMode(long dataLocalMode) {
				set("DataLocalMode", new Variant(dataLocalMode));
			}
			
			public double getPointValue() {
				return get("PointValue").getDouble();
			}
			
			public void setPointValue(double pointValue) {
				set("PointValue", new Variant(pointValue));
			}
			
			public double getMarginDeposit() {
				return get("MarginDeposit").getDouble();
			}
			
			public void setMarginDeposit(double marginDeposit) {
				set("MarginDeposit", new Variant(marginDeposit));
			}			
			
			public double getRoundLotSize() {
				return get("RoundLotSize").getDouble();
			}
			
			public void setRoundLotSize(double roundLotSize) {
				set("RoundLotSize", new Variant(roundLotSize));
			}	
			
			public double getTickSize() {
				return get("TickSize").getDouble();
			}
			
			public void setTickSize(double tickSize) {
				set("TickSize", new Variant(tickSize));
			}
			
			public String getCurrency() {
				return get("Currency").getString();
			}
			
			public void setCurrency(String currency) {
				set("Currency", new Variant(currency));
			}
			
			public String getLastSplitFactor() {
				return get("LastSplitFactor").getString();
			}
			
			public void setLastSplitFactor(String lastSplitFactor) {
				set("LastSplitFactor", new Variant(lastSplitFactor));
			}	
			
			public Date getLastSplitDate() {
				return get("LastSplitDate").getJavaDate();
			}
			
			public void setLastSplitFactor(Date lastSplitDate) {
				set("LastSplitDate", new Variant(lastSplitDate));
			}
			
			public double getDividendPerShare() {
				return get("DividendPerShare").getDouble();
			}
			
			public void setDividendPerShare(double dividendPerShare) {
				set("DividendPerShare", new Variant(dividendPerShare));
			}
			
			public Date getDividendPayDate() {
				return get("DividendPayDate").getJavaDate();
			}
			
			public void setDividendPayDate(double dividendPayDate) {
				set("DividendPayDate", new Variant(dividendPayDate));
			}
			
			public Date getExDividendDate() {
				return get("ExDividendDate").getJavaDate();
			}
			
			public void setExDividendDate(Date exDividendDate) {
				set("ExDividendDate", new Variant(exDividendDate));
			}
			
			public double getPEGRatio() {
				return get("PEGRatio").getDouble();
			}
			
			public void setPEGRatio(double pegRatio) {
				set("PEGRatio", new Variant(pegRatio));
			}
			
			public double getProfitMargin() {
				return get("ProfitMargin").getDouble();
			}
			
			public void setProfitMargin(double profitMargin) {
				set("ProfitMargin", new Variant(profitMargin));
			}
			
			public double getOperatingMargin() {
				return get("OperatingMargin").getDouble();
			}
			
			public void setOperatingMargin(double operatingMargin) {
				set("OperatingMargin", new Variant(operatingMargin));
			}	
			
			public double getOneYearTargetPrice() {
				return get("OneYearTargetPrice").getDouble();
			}
			
			public void setOneYearTargetPrice(double oneYearTargetPrice) {
				set("OneYearTargetPrice", new Variant(oneYearTargetPrice));
			}			
			
			public double getReturnOnAssets() {
				return get("ReturnOnAssets").getDouble();
			}
			
			public void setReturnOnAssets(double returnOnAssets) {
				set("ReturnOnAssets", new Variant(returnOnAssets));
			}	
			
			public double getReturnOnEquity() {
				return get("ReturnOnEquity").getDouble();
			}
			
			public void setReturnOnEquity(double returnOnEquity) {
				set("ReturnOnEquity", new Variant(returnOnEquity));
			}
			
			public double getQtrlyRevenueGrowth() {
				return get("QtrlyRevenueGrowth").getDouble();
			}
			
			public void setQtrlyRevenueGrowth(double qtrlyRevenueGrowth) {
				set("QtrlyRevenueGrowth", new Variant(qtrlyRevenueGrowth));
			}
			
			public double getGrossProfitPerShare() {
				return get("GrossProfitPerShare").getDouble();
			}
			
			public void setGrossProfitPerShare(double grossProfitPerShare) {
				set("GrossProfitPerShare", new Variant(grossProfitPerShare));
			}
			
			public double getSalesPerShare() {
				return get("SalesPerShare").getDouble();
			}
			
			public void setSalesPerShare(double salesPerShare) {
				set("SalesPerShare", new Variant(salesPerShare));
			}
			
			public double getEBITDAPerShare() {
				return get("EBITDAPerShare").getDouble();
			}
			
			public void setEBITDAPerShare(double ebitdaPerShare) {
				set("EBITDAPerShare", new Variant(ebitdaPerShare));
			}	
			
			public double getQtrlyEarningsGrowth() {
				return get("QtrlyEarningsGrowth").getDouble();
			}
			
			public void setQtrlyEarningsGrowth(double qtrlyEarningsGrowth) {
				set("QtrlyEarningsGrowth", new Variant(qtrlyEarningsGrowth));
			}	
			
			public double getInsiderHoldPercent() {
				return get("InsiderHoldPercent").getDouble();
			}
			
			public void setInsiderHoldPercent(double insiderHoldPercent) {
				set("InsiderHoldPercent", new Variant(insiderHoldPercent));
			}	
			
			public double getInstitutionHoldPercent() {
				return get("InstitutionHoldPercent").getDouble();
			}
			
			public void setInstitutionHoldPercent(double institutionHoldPercent) {
				set("InstitutionHoldPercent", new Variant(institutionHoldPercent));
			}
			
			public double getSharesShort() {
				return get("SharesShort").getDouble();
			}
			
			public void setSharesShort(double sharesShort) {
				set("SharesShort", new Variant(sharesShort));
			}			
			
			public double getSharesShortPrevMonth() {
				return get("SharesShortPrevMonth").getDouble();
			}
			
			public void setSharesShortPrevMonth(double sharesShortPrevMonth) {
				set("SharesShortPrevMonth", new Variant(sharesShortPrevMonth));
			}
			
			public double getForwardDividendPerShare() {
				return get("ForwardDividendPerShare").getDouble();
			}
			
			public void setForwardDividendPerShare(double forwardDividendPerShare) {
				set("ForwardDividendPerShare", new Variant(forwardDividendPerShare));
			}
			
			public double getForwardEPS() {
				return get("ForwardEPS").getDouble();
			}
			
			public void setForwardEPS(double forwardEPS) {
				set("ForwardEPS", new Variant(forwardEPS));
			}
			
			public double getEPS() {
				return get("EPS").getDouble();
			}
			
			public void setEPS(double eps) {
				set("EPS", new Variant(eps));
			}			
			
			public double getEPSEstCurrentYear() {
				return get("EPSEstCurrentYear").getDouble();
			}
			
			public void setEPSEstCurrentYear(double epsEstCurrentYear) {
				set("EPSEstCurrentYear", new Variant(epsEstCurrentYear));
			}
			
			public double getEPSEstNextYear() {
				return get("EPSEstNextYear").getDouble();
			}
			
			public void setEPSEstNextYear(double epsEstNextYear) {
				set("EPSEstNextYear", new Variant(epsEstNextYear));
			}	
			
			public double getEPSEstNextQuarter() {
				return get("EPSEstNextQuarter").getDouble();
			}
			
			public void setEPSEstNextQuarter(double epsEstNextQuarter) {
				set("EPSEstNextQuarter", new Variant(epsEstNextQuarter));
			}		
			
			public double getOperatingCashFlow() {
				return get("OperatingCashFlow").getDouble();
			}
			
			public void setOperatingCashFlow(double operatingCashFlow) {
				set("OperatingCashFlow", new Variant(operatingCashFlow));
			}	
			
			public double getLeveredFreeCashFlow() {
				return get("LeveredFreeCashFlow").getDouble();
			}
			
			public void setLeveredFreeCashFlow(double leveredFreeCashFlow) {
				set("LeveredFreeCashFlow", new Variant(leveredFreeCashFlow));
			}		
			
			public class Quotations {
				private Dispatch dQuotations;
				
				private Quotations(Dispatch dQuotations) {
					this.dQuotations = dQuotations;
				}
				
				public Quotation add(Date date) {
					Variant vQuotation = Dispatch.call(dQuotations, "Add", new Variant(date));
					
					return !vQuotation.isNull() ? new Quotation(vQuotation.toDispatch()) : null;
				}
				
				public Quotation item(Date date) {
					Variant vQuotation = Dispatch.call(dQuotations, "Item", new Variant(date));
					
					return !vQuotation.isNull() ? new Quotation(vQuotation.toDispatch()) : null;
				}	
				
				public Quotation item(int index) {
					Variant vQuotation = Dispatch.call(dQuotations, "Item", new Variant(index));
					
					return !vQuotation.isNull() ? new Quotation(vQuotation.toDispatch()) : null;
				}	
				
				public boolean remove(Date date) {
					Variant vResult = Dispatch.call(dQuotations, "Remove",
							new Variant(date));
					return vResult.getBoolean();
				}	
				
				public boolean remove(int index) {
					Variant vResult = Dispatch.call(dQuotations, "Remove",
							new Variant(index));
					return vResult.getBoolean();
				}				
				
				public int count() {
					Variant vResult = Dispatch.get(dQuotations, "Count");
					
					return vResult.getInt();
				}				
				
				public class Quotation {
					private Dispatch dQuotation;
					
					private Quotation(Dispatch dQuotation) {
						this.dQuotation = dQuotation;
					}
					
					private Variant get(String propertyName) {
						return Dispatch.get(dQuotation, propertyName);
					}
					
					private void set(String propertyName, Variant propertyValue) {
						Dispatch.put(dQuotation, propertyName, propertyValue);
					}
					
					public Date getDate() {
						return get("Date").getJavaDate();
					}
					
					public void setDate(Date date) {
						set("Date", new Variant(date));
					}
					
					public float getClose() {
						return get("Close").getFloat();
					}
					
					public void setClose(float close) {
						set("Close", new Variant(close));
					}
					
					public float getOpen() {
						return get("Open").getFloat();
					}
					
					public void setOpen(float open) {
						set("Open", new Variant(open));
					}		
					
					public float getHigh() {
						return get("High").getFloat();
					}
					
					public void setHigh(float high) {
						set("High", new Variant(high));
					}
					
					public float getLow() {
						return get("Low").getFloat();
					}
					
					public void setLow(float low) {
						set("Low", new Variant(low));
					}	
					
					public float getVolume() {
						return get("Volume").getFloat();
					}
					
					public void setVolume(float volume) {
						set("Volume", new Variant(volume));
					}
					
					public float getOpenInt() {
						return get("OpenInt").getFloat();
					}
					
					public void setOpenInt(float openInt) {
						set("OpenInt", new Variant(openInt));
					}	
				}
			}
		}
	}
	
	public class AnalysisDocs {
		private Dispatch dAnalysisDocs;
		
		private AnalysisDocs(Dispatch dAnalysisDocs) {
			this.dAnalysisDocs = dAnalysisDocs;
		}
		
		public AnalysisDoc add() {
			Variant dAnalysisDoc = Dispatch.call(dAnalysisDocs, "Add");
			
			return !dAnalysisDoc.isNull() ? new AnalysisDoc(dAnalysisDoc.toDispatch()) : null;
		}
		
		public void close() {
			Dispatch.callSub(dAnalysisDocs, "Close");
		}
		
		public AnalysisDoc open(String fileName) {
			Variant vResult = Dispatch.call(dAnalysisDocs, "Open", 
					new Variant(fileName));
			return !vResult.isNull() ? new AnalysisDoc(vResult.toDispatch()) : null;
		}
		
		public AnalysisDoc item(int index) {
			Variant dAnalysisDoc =  Dispatch.call(dAnalysisDocs, "Item", new Variant(index));
			
			return !dAnalysisDoc.isNull() ? new AnalysisDoc(dAnalysisDoc.toDispatch()) : null;
		}
		
		public int count() {
			Variant vResult = Dispatch.get(dAnalysisDocs, "Count");
			
			return vResult.getInt();
		}
		
		public class AnalysisDoc {
			private Dispatch dAnalysisDoc;
			
			private static final int ACTION_SCAN = 0;
			private static final int ACTION_EXPLORATION = 0;
			private static final int ACTION_PORTFOLIO_BACKTEST = 0;
			private static final int ACTION_INDIVIDUAL_BACKTEST = 0;
			private static final int ACTION_PORTFOLIO_OPTIMIZATION = 0;
			private static final int ACTION_INDIVIDUAL_OPTIMIZATION = 0;
			private static final int ACTION_WALK_FORWARD_TEST = 0;
			
			
			private AnalysisDoc(Dispatch dAnalysisDoc) {
				this.dAnalysisDoc = dAnalysisDoc;
			}
			
			public boolean isBusy() {
				Variant vResult = Dispatch.get(dAnalysisDoc, "IsBusy");
				
				return vResult.getBoolean();
			}
			
			public void close() {
				Dispatch.callSub(dAnalysisDoc, "Close");
			}
			
			public long export(String fileName) {
				Variant vResult = Dispatch.call(dAnalysisDoc, "Export", 
						new Variant(fileName));
				
				return vResult.getLong();
			}
			
			public long run(int action) {
				Variant vResult = Dispatch.call(dAnalysisDoc, "Export", 
						new Variant(action));
				
				return vResult.getLong();
			}
			
		}
	}
}
