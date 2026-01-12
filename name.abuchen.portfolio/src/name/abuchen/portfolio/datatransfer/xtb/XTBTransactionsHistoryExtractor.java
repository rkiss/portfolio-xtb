package name.abuchen.portfolio.datatransfer.xtb;

import java.io.IOException;
import java.math.BigDecimal;
import java.text.MessageFormat;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import name.abuchen.portfolio.Messages;
import name.abuchen.portfolio.datatransfer.Extractor;
import name.abuchen.portfolio.datatransfer.SecurityCache;
import name.abuchen.portfolio.model.Account;
import name.abuchen.portfolio.model.AccountTransaction;
import name.abuchen.portfolio.model.AccountTransaction.Type;
import name.abuchen.portfolio.model.AttributeType;
import name.abuchen.portfolio.model.BuySellEntry;
import name.abuchen.portfolio.model.Client;
import name.abuchen.portfolio.model.Portfolio;
import name.abuchen.portfolio.model.PortfolioTransaction;
import name.abuchen.portfolio.model.Security;
import name.abuchen.portfolio.model.Transaction.Unit;
import name.abuchen.portfolio.money.ExchangeRate;
import name.abuchen.portfolio.money.ExchangeRateProviderFactory;
import name.abuchen.portfolio.money.ExchangeRateTimeSeries;
import name.abuchen.portfolio.money.Money;
import name.abuchen.portfolio.money.Values;

public class XTBTransactionsHistoryExtractor implements Extractor
{

    private final Client client;
    private final ExchangeRateProviderFactory exchangeRateProviderFactory;
    private Portfolio portfolio = null;
    private Account account = null;
    private SecurityCache securityCache;
    private AttributeType externalIdAttr;
    
    private static record StatementMetadata(String name, String currency, String accountNr)
    {
    }

    private static record XTBCashOperation(String ID, LocalDateTime date, String note, Security security, Money amount,
                    double dblAmount)
    {

    }

    private static final class MissingSecurityException extends RuntimeException
    {
        public MissingSecurityException(String message)
        {
            super(message);
        }
    }

    public XTBTransactionsHistoryExtractor(Client client)
    {
        this.client = client;
        this.exchangeRateProviderFactory = new ExchangeRateProviderFactory(client);
        // try to locate the "externalId" account attribute
        Optional<AttributeType> attr = client.getSettings().getAttributeTypes()
                        .filter(p -> p.getTarget() == Account.class && p.getType() == String.class
                                        && p.getName().equals("externalId")) //$NON-NLS-1$
                        .findFirst();
        if (attr.isPresent())
        {
            externalIdAttr = attr.get();
        }
    }

    public Client getClient()
    {
        return client;
    }

    public Account getAccount()
    {
        return account;
    }

    public Portfolio getPortfolio()
    {
        return portfolio;
    }

    @Override
    public String getLabel()
    {
        return Messages.XTBExtractorLabel;
    }

    @Override
    public List<Item> extract(SecurityCache securityCache, InputFile file, List<Exception> errors)
    {
        this.securityCache = securityCache;
        if (externalIdAttr == null)
        {
            errors.add(new IllegalStateException(
                            "There is no 'externalId' attribute defined for deposit accounts. Define one of type string and set it to the ID of the corresponding XTB account."));
        }
        try (OPCPackage pkg = OPCPackage.open(file.getFile(), PackageAccess.READ))
        {
            try (Workbook wbook = new XSSFWorkbook(pkg))
            {
                List<Item> result = new ArrayList<>();
                if (!isValidDocument(wbook, errors))
                    return Collections.emptyList();

                StatementMetadata metadata = getStatementMetadata(wbook, errors);
                if (metadata == null)
                    return Collections.emptyList();

                locateAccounts(metadata, errors);

                result.addAll(processCashOperations(wbook, metadata, errors));

                result.addAll(processClosedPositions(wbook, metadata, errors));

                System.err.println(String.format("Returned %s entries", result.size()));
                Collections.sort(result, (a, b) -> a.getDate().compareTo(b.getDate()));
                return result;
            }
        }
        catch (InvalidFormatException | IOException e)
        {
            errors.add(e);
            return Collections.emptyList();
        }
    }

    private void locateAccounts(StatementMetadata metadata, List<Exception> errors)
    {
        
        List<Portfolio> portfolios = client.getActivePortfolios().stream()
                        .filter(p -> p.getReferenceAccount().getCurrencyCode().equals(metadata.currency())).toList();
        if (portfolios.isEmpty())
            errors.add(new IllegalStateException(
                            String.format("Cannot locate portfolio for currency %s", metadata.currency())));
        portfolio = portfolios.getFirst();
        account = portfolio.getReferenceAccount();
        if (externalIdAttr != null)
        {
            // try to locate a portfolio with the right externalId
            portfolios = portfolios.stream()
                            .filter(p -> metadata.accountNr()
                                            .equals(p.getReferenceAccount().getAttributes().get(externalIdAttr)))
                            .toList();
            if (portfolios.isEmpty())
                errors.add(new IllegalStateException(
                                String.format("Cannot locate portfolio for XTB account %s, neet to set the externalId attribute to the XTB account nr",
                                                metadata.accountNr())));
            else
            {
                portfolio = portfolios.getFirst();
                account = portfolio.getReferenceAccount();
            }

        }
    }

    private Collection<? extends Item> processCashOperations(Workbook wbook, StatementMetadata metadata, List<Exception> errors)
    {
        List<Item> result = new ArrayList<>();
        Sheet sheet = wbook.getSheet("CASH OPERATION HISTORY");
        int rowNr = 11;
        do
        {
            Row row = sheet.getRow(rowNr);
            if (row == null)
                break;
            long xtbPosition = PoiXSSFUtils.getCellValueAsLong(row, "B", -1, null);
            if (xtbPosition == -1)
            {
                break;
            }
            int lastNrOfErrors = errors.size();
            String xtbOpType = PoiXSSFUtils.getCellValueAsString(row, "C", "", errors);
            LocalDateTime xtbOpTime = PoiXSSFUtils.getCellValueAsLocalDateTime(row, "D", null, errors);
            String xtbNote = PoiXSSFUtils.getCellValueAsString(row, "E", "", errors);
            String xtbSymbol = PoiXSSFUtils.getCellValueAsString(row, "F", "", null);
            double xtbAmount = PoiXSSFUtils.getCellValueAsDouble(row, "G", -1, errors);

            if (lastNrOfErrors == errors.size())
            {

                Security security = null;
                if (!xtbSymbol.isEmpty())
                    security = getSecurity(xtbSymbol, errors);
                Money amount = Money.of(metadata.currency(), Math.round(xtbAmount * Values.Amount.factor()));
                XTBCashOperation xtbCashOp = new XTBCashOperation(String.valueOf(xtbPosition), xtbOpTime, xtbNote,
                                security, amount, xtbAmount);
                Item operation = null;
                switch (xtbOpType)
                {
                    case "deposit": //$NON-NLS-1$
                        operation = createCashOperation(Type.DEPOSIT, xtbCashOp, errors);
                        break;
                    case "Free-funds Interest": //$NON-NLS-1$
                        operation = createCashOperation(Type.INTEREST, xtbCashOp, errors);
                        break;
                    case "Free-funds Interest Tax": //$NON-NLS-1$
                        operation = createCashOperation(Type.INTEREST_CHARGE, xtbCashOp, errors);
                        break;
                    case "stock transfer fee": //$NON-NLS-1$
                        if (security != null)
                        {
                            operation = createCashOperation(xtbAmount < 0 ? Type.FEES : Type.FEES_REFUND, xtbCashOp,
                                            errors);
                        }
                        else
                        {
                            errors.add(new IllegalStateException(
                                            String.format("Cash operation %s should have a security", xtbPosition)));
                        }
                        break;
                    case "DIVIDENT":
                        operation = processDivident(xtbCashOp, errors);
                        break;
                    case "Withholding Tax":
                        if (security != null)
                        {
                            operation = createCashOperation(xtbAmount < 0 ? Type.TAXES : Type.TAX_REFUND, xtbCashOp,
                                            errors);
                        }
                        else
                        {
                            errors.add(new IllegalStateException(
                                            String.format("Cash operation %s should have a security", xtbPosition)));
                        }
                        break;
                    case "Stock purchase":
                        if (security != null)
                        {
                            operation = processStockBuy(xtbCashOp, metadata, errors);
                        }
                        else
                        {
                            errors.add(new IllegalStateException(
                                            String.format("Cash operation %s should have a security", xtbPosition)));
                        }
                        break;
                    case "tax RO":
                        operation = processTaxes(xtbCashOp, metadata, errors);
                        break;
                    case "Stock sale", "close trade": //$NON-NLS-1$//$NON-NLS-2$
                        // we do not process this here, they are processed by
                        // the processClosedPositions method
                        break;
                    default:
                        errors.add(new IllegalStateException(
                                        String.format("Unsupported operation type %s for position %s", xtbOpType,
                                                        xtbPosition)));

                        break;
                }
                if (operation != null)
                {
                    result.add(operation);
                }
            }

            rowNr++;
        }
        while (true);
        return result;
    }

    private Item processTaxes(XTBCashOperation xtbCashOp, StatementMetadata metadata, List<Exception> errors)
    {
        // Romania Tax 3% BAVA.DK
        // RO tax VETH.DE 2025-08-14 (31.00 RON) ///OMI/1102002336/31.00//
        Matcher matcher = Pattern.compile("^Romania Tax\\s+(\\d+(\\.\\d+)?\\s*%)\\s*([A-Z0-9]+(\\.[A-Z0-9]+)?)$")
                        .matcher(xtbCashOp.note());
        int grpIndex = 3;
        if (!matcher.matches())
        {
            // try with the second matcher
            matcher = Pattern.compile("^RO tax\\s+([A-Z0-9]+(\\.[A-Z0-9]+)?)\\s+.+$").matcher(xtbCashOp.note());
            grpIndex = 1;
            if (!matcher.matches())
            {
                errors.add(new IllegalStateException(
                                String.format("Pattern for transaction %s does not match a tax expression, got %s",
                                                xtbCashOp.ID, xtbCashOp.note())));
                return null;
            }
        }
        String sSecurity = matcher.group(grpIndex);
        Security security = getSecurity(sSecurity, errors);
        XTBCashOperation adjCashOp = new XTBCashOperation(xtbCashOp.ID(), xtbCashOp.date(), xtbCashOp.note(), security,
                        xtbCashOp.amount(), xtbCashOp.dblAmount());
        Item result = createCashOperation(xtbCashOp.dblAmount() < 0 ? Type.TAXES : Type.TAX_REFUND, adjCashOp, errors);
        if (!metadata.currency().equals(security.getCurrencyCode()))
        {
            // need to add exchange rate
            ExchangeRateTimeSeries exchangeRateTimeSeries = exchangeRateProviderFactory
                            .getTimeSeries(metadata.currency(), security.getCurrencyCode());
            Optional<ExchangeRate> rateOpt = exchangeRateTimeSeries.lookupRate(adjCashOp.date);
            BigDecimal exchangeRate = BigDecimal.ONE;
            if (rateOpt.isPresent())
            {
                exchangeRate = rateOpt.get().getValue();
            }
            BigDecimal dblSecAmount = BigDecimal.valueOf(Math.abs(adjCashOp.dblAmount())).divide(exchangeRate,
                            Values.MC);
            Money secAmount = Money.of(security.getCurrencyCode(),
                            dblSecAmount.multiply(BigDecimal.valueOf(Values.Amount.factor())).longValue());

            Unit unit = new Unit(Unit.Type.GROSS_VALUE, adjCashOp.amount().absolute(), secAmount,
                            exchangeRate);
            AccountTransaction transaction = (AccountTransaction) result.getSubject();
            transaction.addUnit(unit);
        }
        return result;
    }

    private Item processDivident(XTBCashOperation xtbCashOp, List<Exception> errors)
    {
        // SXXPIEX.DE EUR 0.1696/ SHR
        Matcher matcher = Pattern.compile("^.+\\s+(\\d+(\\.\\d+)?)\\s*/\\s*SHR$").matcher(xtbCashOp.note());
        if (!matcher.matches())
        {
            errors.add(new IllegalStateException(
                            String.format("Pattern for transaction %s does not match a divident expression, got %s",
                                            xtbCashOp.ID, xtbCashOp.note())));
            return null;
        }
        String sDivPerShare = matcher.group(1);
        long divPerShare = 0;
        try
        {
            double dShares = Double.parseDouble(sDivPerShare);
            divPerShare = Math.round(dShares * Values.Share.factor());
        }
        catch (NumberFormatException e)
        {
            errors.add(new IllegalStateException(String.format("Error while procesing transaction %s", xtbCashOp.ID()),
                            e));
            return null;
        }
        Item result = createCashOperation(Type.DIVIDENDS, xtbCashOp, errors);

        AccountTransaction transaction = (AccountTransaction) result.getSubject();

        transaction.setShares(divPerShare);

        return result;
    }

    private Item processStockBuy(XTBCashOperation xtbCashOp, StatementMetadata metadata, List<Exception> errors)
    {
        // need to extract nr of share and price per share from note
        // OPEN BUY 0.6177/26.6177 @ 74.960
        // OPEN BUY 8 @ 10.618
        Matcher matcher = Pattern.compile("^OPEN BUY\\s+(\\d+(\\.\\d+)?)\\s*(/\\d+(\\.\\d+)?)?\\s+@\\s+(\\d+(\\.\\d+)?)$").matcher(xtbCashOp.note());
        if (!matcher.matches())
        {
            errors.add(new IllegalStateException(String.format("Pattern for transaction %s does not match a buy expression, got %s", xtbCashOp.ID, xtbCashOp.note())));
            return null;
        }
        String sShares = matcher.group(1);
        double shares = 0;
        try
        {
            shares = Double.valueOf(sShares);
        }
        catch (NumberFormatException e)
        {
            errors.add(new IllegalStateException(String.format("Error while procesing transaction %s", xtbCashOp.ID()),
                            e));
            return null;
        }
        String sPricePerShare = matcher.group(5);
        double pricePerShare;
        try
        {
            pricePerShare = Double.valueOf(sPricePerShare);
        }
        catch (NumberFormatException e)
        {
            errors.add(new IllegalStateException(String.format("Error while procesing transaction %s", xtbCashOp.ID()),
                            e));
            return null;
        }

        BuySellEntry t = new BuySellEntry(portfolio, account);
        t.setType(PortfolioTransaction.Type.BUY);
        t.setSource(xtbCashOp.ID());
        t.setMonetaryAmount(xtbCashOp.amount().absolute());
        t.setDate(xtbCashOp.date());
        t.setNote(xtbCashOp.note());
        t.setSecurity(xtbCashOp.security());
        t.setShares(Math.round(shares * Values.Share.factor()));
        if (!xtbCashOp.security().getCurrencyCode().equals(metadata.currency()))
        {
            // need to set the forex data
            double grossValueInSecurityCurrency = shares * pricePerShare;
            Money sourceAmount = Money.of(xtbCashOp.security().getCurrencyCode(),
                            Math.round(grossValueInSecurityCurrency * Values.Amount.factor()));

            Unit unit = new Unit(Unit.Type.GROSS_VALUE, xtbCashOp.amount().absolute(), sourceAmount,
                            BigDecimal.valueOf(Math.abs(xtbCashOp.dblAmount()))
                            .divide(BigDecimal.valueOf(grossValueInSecurityCurrency), Values.MC));
            t.getPortfolioTransaction().addUnit(unit);
        }

        BuySellEntryItem item = new BuySellEntryItem(t);

        item.setAccountPrimary(account);
        item.setPortfolioPrimary(portfolio);

        return item;
    }

    private Item createCashOperation(Type type, XTBCashOperation xtbCashOp, List<Exception> errors)
    {
        AccountTransaction t = new AccountTransaction();
        t.setSource(xtbCashOp.ID());
        t.setType(type);
        t.setAmount(Math.abs(xtbCashOp.amount().getAmount()));
        t.setCurrencyCode(xtbCashOp.amount().getCurrencyCode());
        t.setDateTime(xtbCashOp.date());
        t.setNote(xtbCashOp.note());
        if (xtbCashOp.security() != null)
        {
            t.setSecurity(xtbCashOp.security());
        }
        if (type == Type.INTEREST_CHARGE)
        {
            if (xtbCashOp.amount().isPositive())
            {
                // take this as a correction
                t.setType(Type.TAX_REFUND);
            }
            t.addUnit(new Unit(Unit.Type.TAX, Money.of(t.getCurrencyCode(), Math.abs(xtbCashOp.amount().getAmount()))));
        }
        TransactionItem item = new TransactionItem(t);

        item.setAccountPrimary(account);
        item.setPortfolioPrimary(portfolio);

        return item;
    }

    @SuppressWarnings("nls")
    private List<Item> processClosedPositions(Workbook wbook, StatementMetadata metadata, List<Exception> errors)
    {
        List<Item> result = new ArrayList<>();
        Sheet sheet = wbook.getSheet("CLOSED POSITION HISTORY");
        int rowNr = 13;
        do
        {
            Row row = sheet.getRow(rowNr);
            if (row == null)
                break;
            long xtbPosition = PoiXSSFUtils.getCellValueAsLong(row, "B", -1, null);
            if (xtbPosition == -1)
            {
                break;
            }
            String xtbTransactionType = PoiXSSFUtils.getCellValueAsString(row, "D", "", errors);
            if (!xtbTransactionType.isEmpty() && !"BUY".equals(xtbTransactionType))
            {
                errors.add(new IllegalStateException(String.format("Unexpected close type: %s", xtbTransactionType)));
                break;
            }
            int lastNrOfErrors = errors.size();
            String xtbSymbol = PoiXSSFUtils.getCellValueAsString(row, "C", "", errors);
            double xtbVolume = PoiXSSFUtils.getCellValueAsDouble(row, "E", -1, errors);
            // The sale value is in account currency
            double xtbSaleValue = PoiXSSFUtils.getCellValueAsDouble(row, "M", -1, errors);
            double xtbCloseShareValue = PoiXSSFUtils.getCellValueAsDouble(row, "I", -1, errors);
            LocalDateTime xtbCloseTime = PoiXSSFUtils.getCellValueAsLocalDateTime(row, "H", null, errors);

            if (errors.size() == lastNrOfErrors)
            {
                Security security = getSecurity(xtbSymbol, errors);

                if (security != null)
                {

                    var entry = new BuySellEntry();
                    entry.setType(PortfolioTransaction.Type.SELL);
                    entry.setSource(String.valueOf(xtbPosition));
                    entry.setSecurity(security);
                    entry.setDate(xtbCloseTime);
                    Money amount = Money.of(metadata.currency(), Math.round(xtbSaleValue * Values.Amount.factor()));
                    entry.setMonetaryAmount(amount);
                    entry.setNote(String.format("CLOSE BUY %s @ %s", xtbVolume, xtbCloseShareValue));
                    entry.setShares(Math.round(xtbVolume * Values.Share.factor()));

                    if (!security.getCurrencyCode().equals(metadata.currency()))
                    {
                        // need to set the forex data
                        // xtbCloseShareValue
                        double grossValueInSecurityCurrency = xtbCloseShareValue * xtbVolume;
                        Money sourceAmount = Money.of(security.getCurrencyCode(),
                                        Math.round(grossValueInSecurityCurrency * Values.Amount.factor()));

                        Unit unit = new Unit(Unit.Type.GROSS_VALUE, amount, sourceAmount,
                                        BigDecimal.valueOf(xtbSaleValue)
                                                        .divide(BigDecimal.valueOf(grossValueInSecurityCurrency),
                                                                        Values.MC));
                        entry.getPortfolioTransaction().addUnit(unit);
                    }

                    var item = new BuySellEntryItem(entry);

                    item.setAccountPrimary(account);
                    item.setPortfolioPrimary(portfolio);

                    result.add(item);

                }
            }

            rowNr++;
        }
        while (true);
        return result;

    }

    private Security getSecurity(String ticket, List<Exception> errors)
    {
        if (ticket == null || ticket.isEmpty())
        {
            errors.add(new MissingSecurityException("Found empty security ticket"));
            return null;
        }
        String securityName = MessageFormat.format(Messages.CSVImportedSecurityLabel, ticket);
        return securityCache.lookup(null, ticket, null, securityName, () -> {
            errors.add(new MissingSecurityException(String.format("Missing security with ticket '%s'", ticket)));
            return new Security(null, client.getBaseCurrency());
        });
    }

    @SuppressWarnings("nls")
    private boolean isValidDocument(Workbook wbook, List<Exception> errors)
    {
        Sheet sheet = wbook.getSheet("CLOSED POSITION HISTORY");
        if (sheet == null)
        {
            errors.add(new IllegalStateException("Cannot locate sheet with name 'CLOSED POSITION HISTORY'"));
            return false;
        }
        try
        {
            verifyCellContent(sheet, "F", 5, "Name and surname");
            verifyCellContent(sheet, "I", 5, "Account");
            verifyCellContent(sheet, "L", 5, "Currency");
            verifyCellContent(sheet, "B", 9, "CLOSED POSITION HISTORY");
            verifyCellContent(sheet, "B", 12, "Position");
            verifyCellContent(sheet, "C", 12, "Symbol");
            verifyCellContent(sheet, "D", 12, "Type");
            verifyCellContent(sheet, "E", 12, "Volume");
            verifyCellContent(sheet, "F", 12, "Open time");
            verifyCellContent(sheet, "G", 12, "Open price");
            verifyCellContent(sheet, "H", 12, "Close time");
            verifyCellContent(sheet, "I", 12, "Close price");
            verifyCellContent(sheet, "L", 12, "Purchase value");
            verifyCellContent(sheet, "M", 12, "Sale value");
            verifyCellContent(sheet, "T", 12, "Gross P/L");
        }
        catch (IllegalStateException e)
        {
            errors.add(e);
            return false;
        }

        sheet = wbook.getSheet("CASH OPERATION HISTORY");
        if (sheet == null)
        {
            errors.add(new IllegalStateException("Cannot locate sheet with name 'CASH OPERATION HISTORY'"));
            return false;
        }
        try
        {
            verifyCellContent(sheet, "D", 4, "Name and surname");
            verifyCellContent(sheet, "E", 4, "Account");
            verifyCellContent(sheet, "F", 4, "Currency");
            verifyCellContent(sheet, "B", 8, "CASH OPERATION HISTORY");
            verifyCellContent(sheet, "B", 10, "ID");
            verifyCellContent(sheet, "C", 10, "Type");
            verifyCellContent(sheet, "D", 10, "Time");
            verifyCellContent(sheet, "E", 10, "Comment");
            verifyCellContent(sheet, "F", 10, "Symbol");
            verifyCellContent(sheet, "G", 10, "Amount");
        }
        catch (IllegalStateException e)
        {
            errors.add(e);
            return false;
        }
        return true;
    }

    @SuppressWarnings("nls")
    private void verifyCellContent(Sheet sheet, String colIndex, int rowIndex, String expected)
    {
        String sheetName = sheet.getSheetName();

        Row row = sheet.getRow(rowIndex);
        if (row == null)
        {
            throw new IllegalStateException("Row " + rowIndex + " is empty in sheet '" + sheetName + "'");
        }
        String cellValue = PoiXSSFUtils.getCellValueAsString(row, colIndex, "", null);
        try
        {
            if (!cellValue.trim().equals(expected))
            {
                throw new IllegalStateException(
                                "Cell '" + sheetName + "'." + colIndex + rowIndex + " should read '" + expected + "'");
            }
        }
        catch (IllegalStateException e)
        {
            throw new IllegalStateException(
                            "Cell '" + sheetName + "'." + colIndex + rowIndex + " should read '" + expected + "'", e);
        }
    }

    @SuppressWarnings("nls")
    private StatementMetadata getStatementMetadata(Workbook wbook, List<Exception> errors)
    {
        Sheet sheet = wbook.getSheet("CLOSED POSITION HISTORY");
        Row row = sheet.getRow(6);
        String accountOwner = PoiXSSFUtils.getCellValueAsString(row, "F", "", errors);
        String accountNr = PoiXSSFUtils.getCellValueAsString(row, "I", "", errors);
        String currency = PoiXSSFUtils.getCellValueAsString(row, "L", "", errors);
        return new StatementMetadata(accountOwner, currency, accountNr);
    }

}
