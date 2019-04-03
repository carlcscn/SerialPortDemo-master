package com.yang.serialport.ui;

import java.awt.Color;
import java.awt.GraphicsEnvironment;
import java.awt.Point;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.List;

import javax.swing.*;
import javax.swing.event.PopupMenuEvent;
import javax.swing.event.PopupMenuListener;
import javax.swing.filechooser.FileNameExtensionFilter;

import com.yang.serialport.manager.SerialPortManager;
import com.yang.serialport.utils.*;

import gnu.io.PortInUseException;
import gnu.io.SerialPort;

/**
 * 主界面
 * 
 * @author yangle
 */
@SuppressWarnings("all")
public class MainFrame extends JFrame {
	static String path = null;
	static String targetPath = null;
	static String scanGunContent=null;

	// 程序界面宽度
	public final int WIDTH = 530;
	// 程序界面高度
	public final int HEIGHT = 620;

	// 数据显示区
	private JTextArea mDataView = new JTextArea();
	private JScrollPane mScrollDataView = new JScrollPane(mDataView);

	private JTextField pathField = new JTextField();
	private JTextField targetPathField = new JTextField();
	private JTextField numberLocationRow = new JTextField();
	private JTextField numberLocationColumn = new JTextField();
	private JTextField resultLocationRow = new JTextField();
	private JTextField resultLocationColumn = new JTextField();
	private JTextField scanGunField = new JTextField();

	// 串口设置面板
	private JPanel mSerialPortPanel = new JPanel();
	private JLabel mSerialPortLabel = new JLabel("串口");
	private JLabel mBaudrateLabel = new JLabel("波特率");
	private JComboBox mCommChoice = new JComboBox();
	private JComboBox mBaudrateChoice = new JComboBox();
	private ButtonGroup mDataChoice = new ButtonGroup();
	private JRadioButton mDataASCIIChoice = new JRadioButton("ASCII", true);
	private JRadioButton mDataHexChoice = new JRadioButton("Hex");

	// 操作面板
	private JPanel mOperatePanel = new JPanel();
	private JTextArea mDataInput = new JTextArea();
	private JButton mSerialPortOperate = new JButton("打开串口");
	private JButton mSendData = new JButton("发送数据");


	//文本上传面板，添加指定位置如A7
	private JPanel filepanel  = new JPanel();
	private JLabel fileLabel = new JLabel("excel文件路径");
	private JLabel targetFileLabel = new JLabel("目标文件夹路径");
	private JLabel numberLocationLabel = new JLabel("序列号填入位置 行 列 填写数字");
	private JLabel resultLocationLabel = new JLabel("检测结果填入位置 行 列 填写数字");
	private JLabel scanGunLabel = new JLabel("扫码枪输入数据");
	private JButton btn1 = new JButton("浏览");
	private JButton btn2 = new JButton("确定");
	private JButton btn3 = new JButton("浏览");
	private JButton btn4 = new JButton("确定");




	// 串口列表
	private List<String> mCommList = null;
	// 串口对象
	private SerialPort mSerialport;

	public MainFrame() {
		initView();
		initComponents();
		actionListener();
		initData();
	}

	/**
	 * 初始化窗口
	 */
	private void initView() {

		// 关闭程序
		setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
		// 禁止窗口最大化
		setResizable(false);

		// 设置程序窗口居中显示
		Point p = GraphicsEnvironment.getLocalGraphicsEnvironment().getCenterPoint();
		setBounds(p.x - WIDTH / 2, p.y - HEIGHT / 2, WIDTH, HEIGHT);
		this.setLayout(null);
		setTitle("6");
		/*setTitle(scanGun.getscanGunData());*/
	}

	/**
	 * 初始化控件
	 */
	private void initComponents() {
		// 数据显示
		mDataView.setFocusable(false);
		mScrollDataView.setBounds(10, 10, 505, 200);
		add(mScrollDataView);

		// 串口设置
		mSerialPortPanel.setBorder(BorderFactory.createTitledBorder("串口设置"));
		mSerialPortPanel.setBounds(10, 450, 170, 130);//220
		mSerialPortPanel.setLayout(null);
		add(mSerialPortPanel);

		mSerialPortLabel.setForeground(Color.gray);
		mSerialPortLabel.setBounds(10, 25, 40, 20);
		mSerialPortPanel.add(mSerialPortLabel);

		mCommChoice.setFocusable(false);
		mCommChoice.setBounds(60, 25, 100, 20);
		mSerialPortPanel.add(mCommChoice);

		mBaudrateLabel.setForeground(Color.gray);
		mBaudrateLabel.setBounds(10, 60, 40, 20);
		mSerialPortPanel.add(mBaudrateLabel);

		mBaudrateChoice.setFocusable(false);
		mBaudrateChoice.setBounds(60, 60, 100, 20);
		mSerialPortPanel.add(mBaudrateChoice);

		mDataASCIIChoice.setBounds(20, 95, 55, 20);
		mDataHexChoice.setBounds(95, 95, 55, 20);
		mDataChoice.add(mDataASCIIChoice);
		mDataChoice.add(mDataHexChoice);
		mSerialPortPanel.add(mDataASCIIChoice);
		mSerialPortPanel.add(mDataHexChoice);

		//文本上传设置
		filepanel.setBorder(BorderFactory.createTitledBorder("文件上传设置"));
		filepanel.setBounds(10, 220, 505, 225);//220
		filepanel.setLayout(null);
		add(filepanel);
		btn1.setBounds(300, 40, 60, 25);
		btn2.setBounds(370, 40, 60, 25);
		btn1.setVisible(true);
		btn2.setVisible(true);
		filepanel.add(btn1);
		filepanel.add(btn2);
		btn3.setBounds(300, 90, 60, 25);
		btn4.setBounds(370, 90, 60, 25);
		btn3.setVisible(true);
		btn4.setVisible(true);
		filepanel.add(btn3);
		filepanel.add(btn4);


		fileLabel.setBounds(10,10,505,30);
		filepanel.add(fileLabel);

		targetFileLabel.setBounds(10,60,505,30);
		filepanel.add(targetFileLabel);

		numberLocationLabel.setBounds(10,110,505,30);
		filepanel.add(numberLocationLabel);

		resultLocationLabel.setBounds(10,140,505,30);
		filepanel.add(resultLocationLabel);

		scanGunLabel.setBounds(10,170,505,30);
		filepanel.add(scanGunLabel);

		pathField.setBounds(20, 40, 265, 20);
		pathField.setVisible(true);
		filepanel.add(pathField);

		targetPathField.setBounds(20,90,265,20);
		targetPathField.setVisible(true);
		filepanel.add(targetPathField);

		scanGunField.setBounds(20,195,265,20);
		scanGunField.setVisible(true);
		filepanel.add(scanGunField);



		numberLocationRow.setBounds(200, 120, 50, 20);
		numberLocationRow.setVisible(true);
		filepanel.add(numberLocationRow);
		numberLocationColumn.setBounds(260, 120, 50, 20);
		numberLocationColumn.setVisible(true);
		filepanel.add(numberLocationColumn);

		resultLocationRow.setBounds(200, 150, 50, 20);
		resultLocationRow.setVisible(true);
		filepanel.add(resultLocationRow);
		resultLocationColumn.setBounds(260, 150, 50, 20);
		resultLocationColumn.setVisible(true);
		filepanel.add(resultLocationColumn);





		// 操作
		mOperatePanel.setBorder(BorderFactory.createTitledBorder("操作"));
		mOperatePanel.setBounds(200, 450, 315, 130);
		mOperatePanel.setLayout(null);
		add(mOperatePanel);

		mDataInput.setBounds(25, 25, 265, 50);
		mDataInput.setLineWrap(true);
		mDataInput.setWrapStyleWord(true);
		mOperatePanel.add(mDataInput);

		mSerialPortOperate.setFocusable(false);
		mSerialPortOperate.setBounds(45, 95, 90, 20);
		mOperatePanel.add(mSerialPortOperate);

		mSendData.setFocusable(false);
		mSendData.setBounds(180, 95, 90, 20);
		mOperatePanel.add(mSendData);
	}

	/**
	 * 初始化数据
	 */
	private void initData() {
		mCommList = SerialPortManager.findPorts();
		// 检查是否有可用串口，有则加入选项中
		if (mCommList == null || mCommList.size() < 1) {
			ShowUtils.warningMessage("没有搜索到有效串口！");
		} else {
			for (String s : mCommList) {
				mCommChoice.addItem(s);
			}
		}

		mBaudrateChoice.addItem("9600");
		mBaudrateChoice.addItem("19200");
		mBaudrateChoice.addItem("38400");
		mBaudrateChoice.addItem("57600");
		mBaudrateChoice.addItem("115200");
	}

	// 为按钮添加事件




	/**
	 * 按钮监听事件
	 */
	private void actionListener() {
		// 串口
		mCommChoice.addPopupMenuListener(new PopupMenuListener() {

			@Override
			public void popupMenuWillBecomeVisible(PopupMenuEvent e) {
				mCommList = SerialPortManager.findPorts();
				// 检查是否有可用串口，有则加入选项中
				if (mCommList == null || mCommList.size() < 1) {
					ShowUtils.warningMessage("没有搜索到有效串口！");
				} else {
					int index = mCommChoice.getSelectedIndex();
					mCommChoice.removeAllItems();
					for (String s : mCommList) {
						mCommChoice.addItem(s);
					}
					mCommChoice.setSelectedIndex(index);
				}
			}

			@Override
			public void popupMenuWillBecomeInvisible(PopupMenuEvent e) {
				// NO OP
			}

			@Override
			public void popupMenuCanceled(PopupMenuEvent e) {
				// NO OP
			}
		});


		btn1.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(final ActionEvent e) {
				final JFileChooser chooser = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("txt","xls");
				chooser.setFileFilter(filter);
				final int returnVal = chooser.showOpenDialog(btn1);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					final File file = chooser.getSelectedFile();
					path = file.getAbsolutePath();
					pathField.setText(path);
				}
			}
		});
		pathField.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent e) {
				if (new File(pathField.getText()).isDirectory()) {
					path = pathField.getText();
				} else {

					System.out.println("路径名错误");
				}

			}
		});


		btn2.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent e) {
				textFileIO.writePathToText(path);
			}
		});

		btn3.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(final ActionEvent e) {
				final JFileChooser chooser1 = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("txt","xls");
				/*chooser1.setFileFilter(filter);*/
				chooser1.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

				final int returnVal = chooser1.showOpenDialog(btn3);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					final File file = chooser1.getSelectedFile();
					targetPath = file.getAbsolutePath();
					targetPathField.setText(targetPath);
				}
			}
		});
		pathField.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent e) {
				if (new File(targetPathField.getText()).isDirectory()) {
					targetPath = targetPathField.getText();
				} else {

					System.out.println("路径名错误");
				}

			}
		});


		btn4.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent e) {
				textFileIO.writeTargetPathToText(targetPath);
			}
		});

		numberLocationRow.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent e) {
				String numberlocation;

					numberlocation = numberLocationRow.getText();
					textFileIO.writeNumLocationRowToText(numberlocation);


			}
		});

		numberLocationColumn.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent e) {
				String numberlocation;

					numberlocation = numberLocationColumn.getText();
					textFileIO.writeNumLocationColumnToText(numberlocation);


			}
		});

		resultLocationRow.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent e) {
				String resultlocation;

					resultlocation = resultLocationRow.getText();
					textFileIO.writeResultLocationRowToText(resultlocation);


			}
		});

		resultLocationColumn.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent e) {
				String resultlocation;

					resultlocation = resultLocationColumn.getText();
					textFileIO.writeResultLocationColumnToText(resultlocation);


			}
		});

		scanGunField.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(final ActionEvent e) {
				String numberlocation;

				String scanGunContent = scanGunField.getText();
				textFileIO.writeScanGunToText(scanGunContent);
				int a = textFileIO.readNumRow();
				int b = textFileIO.readNumColumn();
				try {
					writeExcel.setExcel(path,a-1,b-1,scanGunContent);
				} catch (IOException e1) {
					e1.printStackTrace();
				}



			}
		});

		// 打开|关闭串口

		mSerialPortOperate.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				if ("打开串口".equals(mSerialPortOperate.getText()) && mSerialport == null) {
					openSerialPort(e);
				} else {
					closeSerialPort(e);
				}
			}
		});

		// 发送数据
		mSendData.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				sendData(e);
			}
		});
	}

	/**
	 * 打开串口
	 * 
	 * @param evt
	 *            点击事件
	 */
	private void openSerialPort(java.awt.event.ActionEvent evt) {
		// 获取串口名称
		String commName = (String) mCommChoice.getSelectedItem();
		// 获取波特率，默认为9600
		int baudrate = 9600;
		String bps = (String) mBaudrateChoice.getSelectedItem();
		baudrate = Integer.parseInt(bps);

		// 检查串口名称是否获取正确
		if (commName == null || commName.equals("")) {
			ShowUtils.warningMessage("没有搜索到有效串口！");
		} else {
			try {
				mSerialport = SerialPortManager.openPort(commName, baudrate);
				if (mSerialport != null) {
					mDataView.setText("串口已打开" + "\r\n");
					mSerialPortOperate.setText("关闭串口");
				}
			} catch (PortInUseException e) {
				ShowUtils.warningMessage("串口已被占用！");
			}
		}

		// 添加串口监听
		SerialPortManager.addListener(mSerialport, new SerialPortManager.DataAvailableListener() {

			@Override
			public void dataAvailable() {
				byte[] data = null;
				String a;
				try {
					if (mSerialport == null) {
						ShowUtils.errorMessage("串口对象为空，监听失败！");
					} else {
						// 读取串口数据
						data = SerialPortManager.readFromPort(mSerialport);
						String resultContent = new String(data);
						int c = textFileIO.readResultRow();
						int d = textFileIO.readResultColumn();
						path = textFileIO.readPath();
						targetPath=textFileIO.readTargetPath();
						scanGunContent = textFileIO.readScanGun();

						writeExcel.setExcel(path,c-1,d-1,resultContent);
						copyExcel.copy2(path,targetPath+"/"+scanGunContent+".xls");


						// 以字符串的形式接收数据
						if (mDataASCIIChoice.isSelected()) {
							mDataView.append(new String(data) + "\r\n");

						}

						// 以十六进制的形式接收数据
						if (mDataHexChoice.isSelected()) {
							mDataView.append(ByteUtils.byteArrayToHexString(data) + "\r\n");
						}
					}
				} catch (Exception e) {
					ShowUtils.errorMessage(e.toString());
					// 发生读取错误时显示错误信息后退出系统
					System.exit(0);
				}
			}
		});
	}

	/**
	 * 关闭串口
	 * 
	 * @param evt
	 *            点击事件
	 */
	private void closeSerialPort(java.awt.event.ActionEvent evt) {
		SerialPortManager.closePort(mSerialport);
		mDataView.setText("串口已关闭" + "\r\n");
		mSerialPortOperate.setText("打开串口");
		mSerialport = null;
	}

	/**
	 * 发送数据
	 * 
	 * @param evt
	 *            点击事件
	 */
	private void sendData(java.awt.event.ActionEvent evt) {
		// 待发送数据
		String data = mDataInput.getText().toString();

		if (mSerialport == null) {
			ShowUtils.warningMessage("请先打开串口！");
			return;
		}

		if ("".equals(data) || data == null) {
			ShowUtils.warningMessage("请输入要发送的数据！");
			return;
		}

		// 以字符串的形式发送数据
		if (mDataASCIIChoice.isSelected()) {
			SerialPortManager.sendToPort(mSerialport, data.getBytes());
		}

		// 以十六进制的形式发送数据
		if (mDataHexChoice.isSelected()) {
			SerialPortManager.sendToPort(mSerialport, ByteUtils.hexStr2Byte(data));
		}
	}



	public static void main(String args[]) {
		java.awt.EventQueue.invokeLater(new Runnable() {
			public void run() {
				new MainFrame().setVisible(true);
			}
		});
	}
}