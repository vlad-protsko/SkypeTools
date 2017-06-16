#pragma once

namespace Skype {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::IO;
	using namespace System::Threading;
	using namespace System::Collections::Generic;
	using namespace System::Net;
	using namespace SKYPE4COMLib;

	/// <summary>
	/// —водка дл€ MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// ќсвободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: SKYPE4COMLib::Skype^ skype = gcnew SKYPE4COMLib::SkypeClass();
	private: System::Windows::Forms::StatusStrip^  statusStrip1;
	private: System::Windows::Forms::ToolStripStatusLabel^  toolStripStatusLabel1;
	private: System::Windows::Forms::ToolStripStatusLabel^  toolStripStatusLabel2;
	private: System::Windows::Forms::PictureBox^  pictureBox1;
	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::MenuStrip^  menuStrip1;
	private: System::Windows::Forms::ToolStripMenuItem^  loadContactsToolStripMenuItem;
	private: System::Windows::Forms::GroupBox^  groupBox1;
	private: System::Windows::Forms::ComboBox^  comboBox1;
	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::GroupBox^  groupBox2;
	private: System::Windows::Forms::ListBox^  listBox1;
	private: System::Windows::Forms::CheckBox^  checkBox1;
	private: System::Windows::Forms::Timer^  timer1;
	private: System::Windows::Forms::Label^  label2;
	private: System::Windows::Forms::NumericUpDown^  numericUpDown1;
	private: System::Windows::Forms::ToolStripStatusLabel^  toolStripStatusLabel3;
	private: System::Windows::Forms::ToolStripStatusLabel^  toolStripStatusLabel4;
	private: System::Windows::Forms::TabControl^  tabControl1;
	private: System::Windows::Forms::TabPage^  tabPage1;
	private: System::Windows::Forms::TabPage^  tabPage2;
	private: System::Windows::Forms::Button^  button3;
	private: System::Windows::Forms::Button^  button2;
	private: System::Windows::Forms::SaveFileDialog^  saveAvatarDialog;
	private: System::Windows::Forms::Button^  button4;
	private: System::ComponentModel::IContainer^  components;

			 /// <summary>
		/// ќб€зательна€ переменна€ конструктора.
		/// </summary>


#pragma region Windows Form Designer generated code
		/// <summary>
		/// “ребуемый метод дл€ поддержки конструктора Ч не измен€йте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->components = (gcnew System::ComponentModel::Container());
			this->statusStrip1 = (gcnew System::Windows::Forms::StatusStrip());
			this->toolStripStatusLabel1 = (gcnew System::Windows::Forms::ToolStripStatusLabel());
			this->toolStripStatusLabel2 = (gcnew System::Windows::Forms::ToolStripStatusLabel());
			this->toolStripStatusLabel3 = (gcnew System::Windows::Forms::ToolStripStatusLabel());
			this->toolStripStatusLabel4 = (gcnew System::Windows::Forms::ToolStripStatusLabel());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
			this->loadContactsToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->groupBox1 = (gcnew System::Windows::Forms::GroupBox());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->numericUpDown1 = (gcnew System::Windows::Forms::NumericUpDown());
			this->checkBox1 = (gcnew System::Windows::Forms::CheckBox());
			this->comboBox1 = (gcnew System::Windows::Forms::ComboBox());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->groupBox2 = (gcnew System::Windows::Forms::GroupBox());
			this->listBox1 = (gcnew System::Windows::Forms::ListBox());
			this->timer1 = (gcnew System::Windows::Forms::Timer(this->components));
			this->tabControl1 = (gcnew System::Windows::Forms::TabControl());
			this->tabPage1 = (gcnew System::Windows::Forms::TabPage());
			this->button4 = (gcnew System::Windows::Forms::Button());
			this->button3 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->tabPage2 = (gcnew System::Windows::Forms::TabPage());
			this->saveAvatarDialog = (gcnew System::Windows::Forms::SaveFileDialog());
			this->statusStrip1->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->menuStrip1->SuspendLayout();
			this->groupBox1->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDown1))->BeginInit();
			this->groupBox2->SuspendLayout();
			this->tabControl1->SuspendLayout();
			this->tabPage1->SuspendLayout();
			this->SuspendLayout();
			// 
			// statusStrip1
			// 
			this->statusStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(4) {
				this->toolStripStatusLabel1,
					this->toolStripStatusLabel2, this->toolStripStatusLabel3, this->toolStripStatusLabel4
			});
			this->statusStrip1->Location = System::Drawing::Point(0, 320);
			this->statusStrip1->Name = L"statusStrip1";
			this->statusStrip1->Size = System::Drawing::Size(385, 22);
			this->statusStrip1->SizingGrip = false;
			this->statusStrip1->TabIndex = 0;
			this->statusStrip1->Text = L"statusStrip1";
			// 
			// toolStripStatusLabel1
			// 
			this->toolStripStatusLabel1->Name = L"toolStripStatusLabel1";
			this->toolStripStatusLabel1->Size = System::Drawing::Size(40, 17);
			this->toolStripStatusLabel1->Text = L"Login:";
			// 
			// toolStripStatusLabel2
			// 
			this->toolStripStatusLabel2->Name = L"toolStripStatusLabel2";
			this->toolStripStatusLabel2->Size = System::Drawing::Size(59, 17);
			this->toolStripStatusLabel2->Text = L"Loading...";
			// 
			// toolStripStatusLabel3
			// 
			this->toolStripStatusLabel3->Name = L"toolStripStatusLabel3";
			this->toolStripStatusLabel3->Size = System::Drawing::Size(48, 17);
			this->toolStripStatusLabel3->Text = L"Friends:";
			// 
			// toolStripStatusLabel4
			// 
			this->toolStripStatusLabel4->Name = L"toolStripStatusLabel4";
			this->toolStripStatusLabel4->Size = System::Drawing::Size(13, 17);
			this->toolStripStatusLabel4->Text = L"0";
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackgroundImageLayout = System::Windows::Forms::ImageLayout::Stretch;
			this->pictureBox1->BorderStyle = System::Windows::Forms::BorderStyle::Fixed3D;
			this->pictureBox1->Location = System::Drawing::Point(6, 12);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(128, 128);
			this->pictureBox1->TabIndex = 1;
			this->pictureBox1->TabStop = false;
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(6, 146);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(128, 23);
			this->button1->TabIndex = 2;
			this->button1->Text = L"GetMyAvatar";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// menuStrip1
			// 
			this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->loadContactsToolStripMenuItem });
			this->menuStrip1->Location = System::Drawing::Point(0, 0);
			this->menuStrip1->Name = L"menuStrip1";
			this->menuStrip1->Size = System::Drawing::Size(385, 24);
			this->menuStrip1->TabIndex = 3;
			this->menuStrip1->Text = L"menuStrip1";
			// 
			// loadContactsToolStripMenuItem
			// 
			this->loadContactsToolStripMenuItem->Name = L"loadContactsToolStripMenuItem";
			this->loadContactsToolStripMenuItem->Size = System::Drawing::Size(95, 20);
			this->loadContactsToolStripMenuItem->Text = L"Load Contacts";
			this->loadContactsToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::loadContactsToolStripMenuItem_Click);
			// 
			// groupBox1
			// 
			this->groupBox1->Controls->Add(this->label2);
			this->groupBox1->Controls->Add(this->numericUpDown1);
			this->groupBox1->Controls->Add(this->checkBox1);
			this->groupBox1->Controls->Add(this->comboBox1);
			this->groupBox1->Controls->Add(this->label1);
			this->groupBox1->Location = System::Drawing::Point(140, 6);
			this->groupBox1->Name = L"groupBox1";
			this->groupBox1->Size = System::Drawing::Size(210, 72);
			this->groupBox1->TabIndex = 4;
			this->groupBox1->TabStop = false;
			this->groupBox1->Text = L"Online Status";
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(89, 50);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(45, 13);
			this->label2->TabIndex = 4;
			this->label2->Text = L"Interval:";
			// 
			// numericUpDown1
			// 
			this->numericUpDown1->Location = System::Drawing::Point(140, 46);
			this->numericUpDown1->Maximum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1000, 0, 0, 0 });
			this->numericUpDown1->Minimum = System::Decimal(gcnew cli::array< System::Int32 >(4) { 100, 0, 0, 0 });
			this->numericUpDown1->Name = L"numericUpDown1";
			this->numericUpDown1->Size = System::Drawing::Size(64, 20);
			this->numericUpDown1->TabIndex = 3;
			this->numericUpDown1->Value = System::Decimal(gcnew cli::array< System::Int32 >(4) { 1000, 0, 0, 0 });
			// 
			// checkBox1
			// 
			this->checkBox1->AutoSize = true;
			this->checkBox1->Location = System::Drawing::Point(14, 49);
			this->checkBox1->Name = L"checkBox1";
			this->checkBox1->Size = System::Drawing::Size(63, 17);
			this->checkBox1->TabIndex = 2;
			this->checkBox1->Text = L"Garland";
			this->checkBox1->UseVisualStyleBackColor = true;
			this->checkBox1->CheckedChanged += gcnew System::EventHandler(this, &MyForm::checkBox1_CheckedChanged);
			// 
			// comboBox1
			// 
			this->comboBox1->DropDownStyle = System::Windows::Forms::ComboBoxStyle::DropDownList;
			this->comboBox1->FormattingEnabled = true;
			this->comboBox1->Location = System::Drawing::Point(92, 19);
			this->comboBox1->Name = L"comboBox1";
			this->comboBox1->Size = System::Drawing::Size(111, 21);
			this->comboBox1->TabIndex = 1;
			this->comboBox1->SelectedIndexChanged += gcnew System::EventHandler(this, &MyForm::comboBox1_SelectedIndexChanged);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(11, 22);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(75, 13);
			this->label1->TabIndex = 0;
			this->label1->Text = L"Current status:";
			// 
			// groupBox2
			// 
			this->groupBox2->Controls->Add(this->listBox1);
			this->groupBox2->Location = System::Drawing::Point(140, 84);
			this->groupBox2->Name = L"groupBox2";
			this->groupBox2->Size = System::Drawing::Size(209, 171);
			this->groupBox2->TabIndex = 5;
			this->groupBox2->TabStop = false;
			this->groupBox2->Text = L"Contacts List";
			// 
			// listBox1
			// 
			this->listBox1->FormattingEnabled = true;
			this->listBox1->Location = System::Drawing::Point(6, 19);
			this->listBox1->Name = L"listBox1";
			this->listBox1->Size = System::Drawing::Size(197, 147);
			this->listBox1->TabIndex = 0;
			// 
			// timer1
			// 
			this->timer1->Interval = 1000;
			this->timer1->Tick += gcnew System::EventHandler(this, &MyForm::timer1_Tick);
			// 
			// tabControl1
			// 
			this->tabControl1->Controls->Add(this->tabPage1);
			this->tabControl1->Controls->Add(this->tabPage2);
			this->tabControl1->Location = System::Drawing::Point(12, 27);
			this->tabControl1->Name = L"tabControl1";
			this->tabControl1->SelectedIndex = 0;
			this->tabControl1->Size = System::Drawing::Size(367, 287);
			this->tabControl1->TabIndex = 6;
			// 
			// tabPage1
			// 
			this->tabPage1->Controls->Add(this->button4);
			this->tabPage1->Controls->Add(this->button3);
			this->tabPage1->Controls->Add(this->button2);
			this->tabPage1->Controls->Add(this->pictureBox1);
			this->tabPage1->Controls->Add(this->groupBox2);
			this->tabPage1->Controls->Add(this->button1);
			this->tabPage1->Controls->Add(this->groupBox1);
			this->tabPage1->Location = System::Drawing::Point(4, 22);
			this->tabPage1->Name = L"tabPage1";
			this->tabPage1->Padding = System::Windows::Forms::Padding(3);
			this->tabPage1->Size = System::Drawing::Size(359, 261);
			this->tabPage1->TabIndex = 0;
			this->tabPage1->Text = L"Index";
			this->tabPage1->UseVisualStyleBackColor = true;
			// 
			// button4
			// 
			this->button4->Enabled = false;
			this->button4->Location = System::Drawing::Point(6, 233);
			this->button4->Name = L"button4";
			this->button4->Size = System::Drawing::Size(128, 23);
			this->button4->TabIndex = 8;
			this->button4->Text = L"SetAsMyAvatar";
			this->button4->UseVisualStyleBackColor = true;
			this->button4->Click += gcnew System::EventHandler(this, &MyForm::button4_Click);
			// 
			// button3
			// 
			this->button3->Enabled = false;
			this->button3->Location = System::Drawing::Point(6, 204);
			this->button3->Name = L"button3";
			this->button3->Size = System::Drawing::Size(128, 23);
			this->button3->TabIndex = 7;
			this->button3->Text = L"SaveCurrentAvatar";
			this->button3->UseVisualStyleBackColor = true;
			this->button3->Click += gcnew System::EventHandler(this, &MyForm::button3_Click);
			// 
			// button2
			// 
			this->button2->Location = System::Drawing::Point(6, 175);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(128, 23);
			this->button2->TabIndex = 6;
			this->button2->Text = L"GetUserAvatar";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			// 
			// tabPage2
			// 
			this->tabPage2->Location = System::Drawing::Point(4, 22);
			this->tabPage2->Name = L"tabPage2";
			this->tabPage2->Padding = System::Windows::Forms::Padding(3);
			this->tabPage2->Size = System::Drawing::Size(359, 261);
			this->tabPage2->TabIndex = 1;
			this->tabPage2->Text = L"Next";
			this->tabPage2->UseVisualStyleBackColor = true;
			// 
			// saveAvatarDialog
			// 
			this->saveAvatarDialog->Filter = L"Images (*.jpg)|*.jpg";
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(385, 342);
			this->Controls->Add(this->tabControl1);
			this->Controls->Add(this->statusStrip1);
			this->Controls->Add(this->menuStrip1);
			this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedSingle;
			this->MainMenuStrip = this->menuStrip1;
			this->MaximizeBox = false;
			this->Name = L"MyForm";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = L"SkypeTools";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			this->statusStrip1->ResumeLayout(false);
			this->statusStrip1->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->menuStrip1->ResumeLayout(false);
			this->menuStrip1->PerformLayout();
			this->groupBox1->ResumeLayout(false);
			this->groupBox1->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->numericUpDown1))->EndInit();
			this->groupBox2->ResumeLayout(false);
			this->tabControl1->ResumeLayout(false);
			this->tabPage1->ResumeLayout(false);
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void MyForm_Load(System::Object^  sender, System::EventArgs^  e) {
		/*
		If skype not running, start him.
		Get your online status, add him to combobox and select him
		Subscribe to the event SkypeStatusChange
		*/
		try {
			this->Icon = SystemIcons::Application;
			if (!skype->Client->IsRunning) {
				skype->Client->Start(false, true);
			}

			toolStripStatusLabel2->Text = skype->CurrentUserHandle;
			toolStripStatusLabel4->Text = skype->Friends->Count.ToString();
			for each (String^ OnlineStatus in Enum::GetNames(TUserStatus::typeid))
			{
				comboBox1->Items->Add(OnlineStatus);
			}
			comboBox1->SelectedIndex = (int)Enum::Parse(TUserStatus::typeid, skype->CurrentUserStatus.ToString());
			//Subscribe to the event for cheking skype change status
			this->skype->UserStatus += gcnew SKYPE4COMLib::_ISkypeEvents_UserStatusEventHandler(this, &MyForm::OnUserStatus);
		}
		//For errors
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message);
		}
	}
	private: System::Void OnUserStatus(TUserStatus Status) {
		/*
		If skype change online status, change him in comboBox1
		*/
		int StatusIndex = (int)Enum::Parse(TUserStatus::typeid, Status.ToString());
		comboBox1->SelectedIndex = StatusIndex;
	}
	private: String^ AvatarPath(String^ UserHandle) {
		/*
		Skype api get avatar
		*/
		pictureBox1->SizeMode = PictureBoxSizeMode::StretchImage;
		SKYPE4COMLib::CommandClass^ cmd = gcnew SKYPE4COMLib::CommandClass();
		String^ p = Path::GetTempPath() + UserHandle + ".jpg";
		cmd->Command = "GET USER " + UserHandle + " AVATAR 1 " + p;
		skype->SendCommand(cmd);
		Thread::Sleep(1000); //Sleeping to give time to save
		if (!button3->Enabled || button4->Enabled) {
			button3->Enabled = true;
			button4->Enabled = true;
		}
		return p;
	}
	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {
		/*
		Get your avatar to pictureBox1
		*/
		try {
			pictureBox1->ImageLocation = AvatarPath(skype->CurrentUserHandle);
		}
		//For errors
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void loadContactsToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {
		/*
		Load skype contacts to listBox1
		*/
		loadContactsToolStripMenuItem->Enabled = false;
		listBox1->Items->Clear();
		try {
			for each (User^ SkypeFriend in skype->Friends)
			{
				listBox1->Items->Add(SkypeFriend->Handle);
			}
			MessageBox::Show("Contacts loaded!", "Successfully", MessageBoxButtons::OK, MessageBoxIcon::Information);
		}
		//For errors
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
		loadContactsToolStripMenuItem->Enabled = true;
	}
	private: System::Void comboBox1_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e) {
		try {
			TUserStatus UserStatus = (TUserStatus)Enum::Parse(TUserStatus::typeid, comboBox1->Text);
			skype->ChangeUserStatus(UserStatus);
		}
		//For errors
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void checkBox1_CheckedChanged(System::Object^  sender, System::EventArgs^  e) {
		/*
		Set interval for Garland timer and Enable/Disable him.
		*/
		try {
			timer1->Interval = (int)numericUpDown1->Value;
			if (checkBox1->Checked) {
				timer1->Enabled = true;
			}
			else {
				timer1->Enabled = false;
			}
		}
		//For errors
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void timer1_Tick(System::Object^  sender, System::EventArgs^  e) {
		/*
		Change User Online Status Timer
		*/
		try {
			switch (comboBox1->SelectedIndex)
			{
			case 1: comboBox1->SelectedIndex = 2; break;
			case 2: comboBox1->SelectedIndex = 4; break;
			case 4: comboBox1->SelectedIndex = 5; break;
			case 5: comboBox1->SelectedIndex = 1; break;
			default:
				break;
			}
			comboBox1_SelectedIndexChanged(this, nullptr);
		}
		//For errors
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void button2_Click(System::Object^  sender, System::EventArgs^  e) {
		try {
			if (listBox1->Text == String::Empty) {
				MessageBox::Show("Load/Select contact in list", "Warning", MessageBoxButtons::OK, MessageBoxIcon::Warning);
			}
			else {
				pictureBox1->ImageLocation = AvatarPath(listBox1->Text);
			}
		}
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void button3_Click(System::Object^  sender, System::EventArgs^  e) {
		try {
			saveAvatarDialog->FileName = Path::GetFileName(pictureBox1->ImageLocation);
			if (saveAvatarDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
				File::Copy(pictureBox1->ImageLocation, saveAvatarDialog->FileName);
				MessageBox::Show("Avatar saved to: " + saveAvatarDialog->FileName, "Successfully", MessageBoxButtons::OK, MessageBoxIcon::Information);
			}
		}
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void button4_Click(System::Object^  sender, System::EventArgs^  e) {
		try {
			SKYPE4COMLib::CommandClass^ cmd = gcnew SKYPE4COMLib::CommandClass();
			cmd->Command = "SET AVATAR 1 " + pictureBox1->ImageLocation;
			skype->SendCommand(cmd);
			MessageBox::Show("Avatar installed!", "Successfully", MessageBoxButtons::OK, MessageBoxIcon::Information);
		}
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
};
}
