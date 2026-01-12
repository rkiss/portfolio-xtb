package name.abuchen.portfolio.ui.handlers;

import java.io.File;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jakarta.inject.Named;

import org.eclipse.e4.core.di.annotations.CanExecute;
import org.eclipse.e4.core.di.annotations.Execute;
import org.eclipse.e4.ui.model.application.ui.basic.MPart;
import org.eclipse.e4.ui.services.IServiceConstants;
import org.eclipse.jface.dialogs.Dialog;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.jface.preference.IPreferenceStore;
import org.eclipse.jface.wizard.WizardDialog;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Shell;

import name.abuchen.portfolio.datatransfer.Extractor;
import name.abuchen.portfolio.datatransfer.xtb.XTBTransactionsHistoryExtractor;
import name.abuchen.portfolio.model.Account;
import name.abuchen.portfolio.model.Client;
import name.abuchen.portfolio.model.Portfolio;
import name.abuchen.portfolio.ui.Messages;
import name.abuchen.portfolio.ui.PortfolioPlugin;
import name.abuchen.portfolio.ui.UIConstants;
import name.abuchen.portfolio.ui.editor.FilePathHelper;
import name.abuchen.portfolio.ui.editor.PortfolioPart;
import name.abuchen.portfolio.ui.wizards.datatransfer.ImportExtractedItemsWizard;

public class ImportXTBHandler
{
    @CanExecute
    boolean isVisible(@Named(IServiceConstants.ACTIVE_PART) MPart part)
    {
        return MenuHelper.isClientPartActive(part);
    }

    @Execute
    public void execute(@Named(IServiceConstants.ACTIVE_PART) MPart part,
                    @Named(IServiceConstants.ACTIVE_SHELL) Shell shell)
    {
        doExecute(part, shell);
    }

    /* package */ void doExecute(MPart part, Shell shell)
    {
        MenuHelper.getActiveClient(part)
                        .ifPresent(client -> runImport((PortfolioPart) part.getObject(), shell, client, null, null));
    }

    public static void runImport(PortfolioPart part, Shell shell, Client client, Account account, Portfolio portfolio)
    {
        if (client.getAccounts().isEmpty())
        {
            MessageDialog.openError(shell, Messages.LabelError, Messages.MsgErrorAccountNotExist);
            return;
        }

        if (client.getPortfolios().isEmpty())
        {
            MessageDialog.openError(shell, Messages.LabelError, Messages.MsgErrorPortfolioNotExist);
            return;
        }

        XTBTransactionsHistoryExtractor extractor = new XTBTransactionsHistoryExtractor(client);
        FilePathHelper helper = new FilePathHelper(part, UIConstants.Preferences.DEFAULT_OPEN_PATH);

        FileDialog fileDialog = new FileDialog(shell, SWT.OPEN | SWT.MULTI);
        fileDialog.setText(Messages.XTBImportWizardAssistant);
        fileDialog.setFilterNames(new String[] { Messages.XTBImportFilterName });
        fileDialog.setFilterExtensions(new String[] { "*.xlsx" }); //$NON-NLS-1$
        fileDialog.setFilterPath(helper.getPath());
        fileDialog.open();

        String[] filenames = fileDialog.getFileNames();

        if (filenames.length == 0)
            return;

        helper.savePath(fileDialog.getFilterPath());

        List<Extractor.InputFile> files = new ArrayList<>();
        for (String filename : filenames)
            files.add(new Extractor.InputFile(new File(fileDialog.getFilterPath(), filename)));

        ArrayList<Exception> errors = new ArrayList<>();
        ArrayList<Extractor.Item> results = new ArrayList<>();

        IPreferenceStore preferences = part.getPreferenceStore();

        try
        {
            IRunnableWithProgress operation = monitor -> {
                results.addAll(extractor.extract(files, errors));
            };

            new ProgressMonitorDialog(shell).run(true, true, operation);

            Map<File, List<Exception>> e = new HashMap<>();
            if (!errors.isEmpty())
                e.put(files.get(0).getFile(), errors);


            Display.getDefault().asyncExec(() -> {
                ImportExtractedItemsWizard wizard = new ImportExtractedItemsWizard(client, preferences,
                                Collections.singletonMap(extractor, results), e);
                wizard.setTarget(extractor.getAccount());
                wizard.setTarget(extractor.getPortfolio());
                Dialog wizardDialog = new WizardDialog(shell, wizard);
                wizardDialog.open();
            });
        }
        catch (InvocationTargetException | InterruptedException e)
        {
            PortfolioPlugin.log(e);
            String message = e.getCause() != null ? e.getCause().getMessage() : e.getMessage();
            MessageDialog.openError(shell, Messages.LabelError, message);
        }
        
        
    }

}
