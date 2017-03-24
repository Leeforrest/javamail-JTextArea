package xiaodong.mail_4_linlin;

import java.awt.HeadlessException;
import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileSystemView;

public class FileChooser {
	File selectedFile;

	public FileChooser(String title) {
		FileSystemView view = FileSystemView.getFileSystemView();
		JFileChooser fileChooser = new JFileChooser(view);
		fileChooser.setMultiSelectionEnabled(false);
		fileChooser.setDialogTitle(title);
		javax.swing.filechooser.FileFilter filter = new DBFilter();
		fileChooser.setFileFilter(filter);
		fileChooser.setCurrentDirectory(view.getDefaultDirectory());
		try {
			fileChooser.showOpenDialog(null);
		} catch (HeadlessException ex) {
			System.out.println("Keyboard and Mouse Required");
			ex.printStackTrace();
			System.exit(-1);
		}

		selectedFile = fileChooser.getSelectedFile();
		if (selectedFile == null) {
			System.out.println("Database not selected");
			System.exit(-1);
		}

	}

	public File getSelectedFile() {
		return selectedFile;
	}
}


class DBFilter extends javax.swing.filechooser.FileFilter {
	String description = "Excel File";

	/**
	 * Whether the given file is accepted by this filter.
	 */
	@Override
	public boolean accept(File f) {
		// We need to allow the user to be able to navigate directories
		if (f.isDirectory()) {
			return true;
		}

		// Only allow the user to select a file with a "db" extension
		String name = f.getName().toLowerCase();
		if (name.endsWith("xlsx") || name.endsWith("pdf")) {
			return true;
		}

		return false;
	}

	/**
	 * The description of the filter.
	 */
	@Override
	public String getDescription() {
		return description;
	}
}