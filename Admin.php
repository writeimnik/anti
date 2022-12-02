<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
class Admin extends CI_Controller {


	public function __construct()
	{
		parent::__construct();
		$this->load->model('Admin_model');
		$this->load->model('Report_model');
	}


	public function index()
	{
		if($this->input->post())
		{
			$email = $this->input->post('email');
			$password = $this->input->post('password');
			$result=$this->db->get_where('admin_account', array('admin_email'=>$email, 'password'=>$password, 'status'=> 1))->result_array();
			@$admin_id = $result[0]['admin_id'];
			@$password = $result[0]['password'];
			@$admin_name = $result[0]['admin_name'];
			$this->session->set_userdata('admin_id', $admin_id);
			$this->session->set_userdata('password', $password);
			$this->session->set_userdata('admin_name', $admin_name);
		

			$query=$this->db->get_where('admin_account', array('admin_email'=>$email, 'password'=>$password, 'status'=> 1));
			
			if($query->num_rows() == 0)
			{
				$page_data['check'] = '<div class="col-lg-12" style="height: 2rem; display: flex;align-items: center;justify-content: center; background: #fadbd8; color: #78261f;">admin ID or password is wrong</div>';
			}

		}
		if(@$this->session->userdata['admin_id'])
		{
			$this->dashboard();
		}
		else
		{
			$page_data['login_title'] = "Admin Login";
			$page_data['page_title'] = "Admin Page";
			$page_data['page'] = "admin/login";
			$this->load->view('admin/layouts/login', $page_data);
		}
	}
	public function login()
	{
		
		if($this->input->post());
		{
			$email = $this->input->post('email');
			$password = $this->input->post('password');
			$result=$this->db->get_where('admin_account', array('admin_email'=>$email, 'password'=>$password, 'status'=> 1))->result_array();
			@$admin_id = $result[0]['admin_id'];
			@$password = $result[0]['password'];
			@$admin_name = $result[0]['admin_name'];
			$this->session->set_userdata('admin_id', $admin_id);
			$this->session->set_userdata('password', $password);
			$this->session->set_userdata('admin_name', $admin_name);
			


			$query=$this->db->get_where('admin_account', array('admin_email'=>$email, 'password'=>$password, 'status'=> 1));
			
			if($query->num_rows() == 0)
			{
				$page_data['check'] = '<div class="col-lg-12" style="height: 2rem; display: flex;align-items: center;justify-content: center; background: #fadbd8; color: #78261f;">admin ID or password is wrong</div>';
			}


		}
		if($this->session->userdata['admin_id'])
		{
			// print_r("expression");
			// die();
			// redirect to dashbord
			$this->dashboard();

		}
		else
		{
			$page_data['login_title'] = "Admin Login";
			$page_data['page_title'] = "Admin Page";
			$page_data['page'] = "admin/login";
			$this->load->view('admin/layouts/login', $page_data);
		}
	}

	public function dashboard()
	{

		if($this->session->userdata['admin_id'])
		{
			$admin_name = $this->session->userdata['admin_name'];	
			$admin_id = $this->session->userdata['admin_id'];			


			$page_data['todaycases'] = $this->Admin_model->todaycases();	
			$page_data['totalcases'] = $this->Admin_model->totalcases();


			$page_data['total_agent'] = $this->Admin_model->total_agent();	
			$page_data['active_agent'] = $this->Admin_model->active_agent();	
			
			$page_data['alcoholsale'] = $this->Admin_model->alcoholsale();	
			$page_data['drugsale'] = $this->Admin_model->drugsale();	
			$page_data['alcoholism'] = $this->Admin_model->alcoholism();	
			$page_data['drugabuse'] = $this->Admin_model->drugabuse();	
			$page_data['proxy'] = $this->Admin_model->proxy_cases();	




			//Redirect to dashbord
			$page_data['admin_name'] =  $admin_name;
			$page_data['admin_id'] =  $admin_id;
			$page_data['page_title'] = "Admin Dashbord";
			$page_data['page'] = "admin/dashboard";
			$this->load->view('admin/layouts/dashboard', $page_data);
		}
		else
		{
			redirect("index.php/Admin/login");		
		}
	}

	public function active_agent()
	{

	
		if($this->session->userdata['admin_id'])
		{
			
			$active_agents_list = $this->Admin_model->active_agents_list();

			// echo "<pre>";
			// print_r($agents_list);
			// die();
			$page_data['admin_name'] =  $this->session->userdata['admin_name'];
			$page_data['agents_list'] = $active_agents_list;
			$page_data['page_title'] = "Active Agent";
			$page_data['page'] = "admin/active_agent";
			$this->load->view('admin/layouts/dashboard', $page_data);
		}
		else
		{
			redirect("index.php/Admin/login");		
		}
	
	}
	public function report_summery()
	{
		if($this->session->userdata['admin_id'])
		{
			$admin_name = $this->session->userdata['admin_name'];	
			$admin_id = $this->session->userdata['admin_id'];			


			// type of caller data
			$page_data['FriendColleague'] = $this->Report_model->FriendColleague();
			$page_data['Familymember'] = $this->Report_model->Familymember();
			$page_data['NotDisclosed'] = $this->Report_model->NotDisclosed();
			$page_data['OnbehalfOforganization'] = $this->Report_model->OnbehalfOforganization();
			$page_data['ProxyPrankCall'] = $this->Report_model->ProxyPrankCall();
			$page_data['self'] = $this->Report_model->self();
			$page_data['Neighbour'] = $this->Report_model->Neighbour();
			$page_data['type_of_caller_Others'] = $this->Report_model->type_of_caller_Others();

		
			$page_data['lessthan15years'] = $this->Report_model->lessthan15years();
			$page_data['age15_20years'] = $this->Report_model->age15_20years();
			$page_data['age21_30years'] = $this->Report_model->age21_30years();
			$page_data['age31_40years'] = $this->Report_model->age31_40years();
			$page_data['age41_50years'] = $this->Report_model->age41_50years();
			$page_data['more_than_50_years'] = $this->Report_model->more_than_50_years();
			$page_data['total_age_data'] = $page_data['lessthan15years']+$page_data['age15_20years']+$page_data['age21_30years']+$page_data['age31_40years']+$page_data['age41_50years']+$page_data['more_than_50_years'];

			// Occupation
			$page_data['LabourFarmers'] = $this->Report_model->LabourFarmers();
			$page_data['Selfemployed'] = $this->Report_model->Selfemployed();
			$page_data['Service'] = $this->Report_model->Service();
			$page_data['Student'] = $this->Report_model->Student();
			$page_data['Unemployed'] = $this->Report_model->Unemployed();
			$page_data['OthersOccupation'] = $this->Report_model->OthersOccupation();
			$page_data['total_Occupation_data'] = $page_data['LabourFarmers']+$page_data['Selfemployed']+$page_data['Service']+$page_data['Student']+$page_data['Unemployed']+$page_data['OthersOccupation'];

			// Source of information
			$page_data['Internet'] = $this->Report_model->Internet();
			$page_data['Newspaper'] = $this->Report_model->Newspaper();
			$page_data['Television'] = $this->Report_model->Television();
			$page_data['AwarenessProgram'] = $this->Report_model->AwarenessProgram();
			$page_data['otherSourceofinformation'] = $this->Report_model->otherSourceofinformation();
			$page_data['total_Sourceofinformation'] = $page_data['Internet']+$page_data['Newspaper']+$page_data['Television']+$page_data['AwarenessProgram']+$page_data['otherSourceofinformation'];

			// age

			// gender
			$page_data['female'] = $this->Report_model->female();
			$page_data['male'] = $this->Report_model->male();
			$page_data['transgender'] = $this->Report_model->transgender();
			$page_data['total_gender'] = $page_data['female']+$page_data['male']+$page_data['transgender'];

			// Frequency_of_consumption1
			$page_data['Occasional_consumption1'] = $this->Report_model->Occasional_consumption1();
			$page_data['Habitual'] = $this->Report_model->Habitual();
			$page_data['total_gender_Frequency_of_consumption1'] = $page_data['Occasional_consumption1']+$page_data['Habitual'];

			// Historyalcoholic
			$page_data['Yes_Historyalcoholic'] = $this->Report_model->Yes_Historyalcoholic();
			$page_data['No_Historyalcoholic'] = $this->Report_model->No_Historyalcoholic();
			$page_data['Total_Historyalcoholic'] = $page_data['Yes_Historyalcoholic']+$page_data['No_Historyalcoholic'];
			
			// ifAlcoProblem_Identified
			$page_data['Physical'] = $this->Report_model->Physical();
			$page_data['Psychological'] = $this->Report_model->Psychological();
			$page_data['Social'] = $this->Report_model->Social();
			$page_data['Suicidal_Tendency'] = $this->Report_model->Suicidal_Tendency();
			$page_data['Others_AlcoProblem_Identified'] = $this->Report_model->Others_AlcoProblem_Identified();
			$page_data['Total_AlcoProblem_Identified'] = $page_data['Physical']+$page_data['Psychological']+$page_data['Social']+$page_data['Suicidal_Tendency']+$page_data['Others_AlcoProblem_Identified'];

			//ClientLIvingStatus
			$page_data['client_living_status'] = $this->Report_model->client_living_status();
			$page_data['client_with_friend_hostal'] = $this->Report_model->client_with_friend_hostal();
			$page_data['client_with_family_relatives'] = $this->Report_model->client_with_family_relatives();
			$page_data['living_alone'] = $this->Report_model->living_alone();
			$page_data['slum_street_public_place'] = $this->Report_model->slum_street_public_place();
			$page_data['Others_ClientLIvingStatus'] = $this->Report_model->Others_ClientLIvingStatus();
			$page_data['Total_ClientLIvingStatus'] = $page_data['client_living_status']+$page_data['client_with_friend_hostal']+$page_data['client_with_family_relatives']+$page_data['living_alone']+$page_data['slum_street_public_place']+$page_data['Others_ClientLIvingStatus'];


			// alcoholism ends here

			// drug abuse

			// Problem Identified Counsellor  
			$page_data['Physical_drugabuse'] = $this->db->get_where('drugabuse',array('IFDrugtypeofCon'=> 'Physical'))->num_rows();
			$page_data['Psychological_drugabuse'] = $this->db->get_where('drugabuse',array('IFDrugtypeofCon'=> 'Psychological'))->num_rows();
			$page_data['Social_drugabuse'] = $this->db->get_where('drugabuse',array('IFDrugtypeofCon'=> 'Social'))->num_rows();
			$page_data['Suicidal_Tendency_drugabuse'] =$this->db->get_where('drugabuse',array('IFDrugtypeofCon'=> 'Suicidal Tendency'))->num_rows();
			$page_data['Others_Identified_drugabuse'] =$this->db->get_where('drugabuse',array('IFDrugtypeofCon'=> 'Others'))->num_rows();
			$page_data['Total_drugabuse'] = $page_data['Physical']+$page_data['Psychological']+$page_data['Social']+$page_data['Suicidal_Tendency']+$page_data['Others_AlcoProblem_Identified'];
			// Injecting Drug User (IDU)

			$page_data['YesInjecting'] = $this->db->get_where('drugabuse',array('injecting_drug'=> 'Yes'))->num_rows();
			$page_data['No_Injecting'] = $this->db->get_where('drugabuse',array('injecting_drug'=> 'No'))->num_rows();
			$page_data['dontknowInjecting'] = $this->db->get_where('drugabuse',array('injecting_drug'=> "Don't Know"))->num_rows();
			
			$page_data['Total_Injecting'] = $page_data['YesInjecting']+$page_data['No_Injecting']+$page_data['dontknowInjecting'];

			//Redirect to dashbord
			$page_data['admin_name'] =  $admin_name;
			$page_data['admin_id'] =  $admin_id;
			$page_data['page_title'] = "Admin Dashbord";
			$page_data['page'] = "admin/report_summery";
			$this->load->view('admin/layouts/dashboard', $page_data);
		}
		else
		{
			redirect("index.php/Admin/login");		
		}
	}

	public function create_agent()
	{

		if($this->session->userdata['admin_id'])
		{
			if($this->input->post())
			{
				$name = $this->input->post('name');
				$gender = $this->input->post('gender');
				$mobile = $this->input->post('mobile');
				$agent_email = $this->input->post('email');
				$password = $this->input->post('password');
				$supervisor_id = $this->input->post('supervisor');
				$supervisor_details = $this->Admin_model->supervisor_details($supervisor_id);


				$supervisor_name = $supervisor_details[0]['supervisor_name'];
				
				$user_data = array
				(
					"agent_name" =>$name,
					"agent_email" =>$agent_email,
					"password" =>$password,
					"gender" =>$gender,
					"supervisor_name" =>$supervisor_name,
					"supervisor_id" =>$supervisor_id,
					"status" =>1,
				);
				$this->Admin_model->insert_user($user_data);
				redirect("index.php/Admin/all_agents");
			}
			else
			{


				$name = $this->session->userdata['admin_name'];	
				$admin_id = $this->session->userdata['admin_id'];
				$all_supervisor = $this->Admin_model->supervisors_list();
				//Redirect to dashbord
				$page_data['admin_name'] =  $name;
				$page_data['all_supervisor'] = $all_supervisor;

				$page_data['page_title'] = "Create Agent";
				$page_data['page'] = "admin/create_agent";
				$this->load->view('admin/layouts/dashboard', $page_data);
			}
			
		}
		else
		{
			redirect("index.php/Admin/login");		
		}
	}
	public function edit_agent($agent_id)
	{
		$all_supervisor = $this->Admin_model->supervisors_list();
		$page_data['all_supervisor'] = $all_supervisor;
		$agent_details = $this->Admin_model->agent_details($agent_id);
		$page_data['agent_details'] = $agent_details;
		$page_data['admin_name'] =  $this->session->userdata['admin_name'];
		$page_data['page'] = "admin/edit_agent";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}
	public function update_agent()
	{
		$agent_id = $this->input->post('agent_id');
		$name = $this->input->post('name');
		$gender = $this->input->post('gender');
		$mobile = $this->input->post('mobile');
		$agent_email = $this->input->post('email');
		$password = $this->input->post('password');
		$supervisor_id = $this->input->post('supervisor');

		$supervisor_details = $this->Admin_model->supervisor_details($supervisor_id);
		$supervisor_name = $supervisor_details[0]['supervisor_name'];
		$agent_data = array
		(
			"agent_name" =>$name,
			"agent_email" =>$agent_email,
			"password" =>$password,
			"gender" =>$gender,
			"supervisor_name" =>$supervisor_name,
			"supervisor_id" =>$supervisor_id,
		);
		$this->Admin_model->edit_agent($agent_data, $agent_id);
		redirect("index.php/Admin/agent_details/".$agent_id);
	}

	public function agent_details($user_id)
	{
		$agent_details = $this->Admin_model->agent_details($user_id);
		$page_data['admin_name'] =  $this->session->userdata['admin_name'];
		$page_data['agent_details'] = $agent_details;
		$page_data['page_title'] = "Create User";
		$page_data['page'] = "admin/agent_details";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}


	public function all_agents()
	{	
		if($this->session->userdata['admin_id'])
		{
			
			$agents_list = $this->Admin_model->agents_list();

			// echo "<pre>";
			// print_r($agents_list);
			// die();
			$page_data['admin_name'] =  $this->session->userdata['admin_name'];
			$page_data['agents_list'] = $agents_list;
			$page_data['page_title'] = "Create User";
			$page_data['page'] = "admin/all_agents";
			$this->load->view('admin/layouts/dashboard', $page_data);
		}
		else
		{
			redirect("index.php/Admin/login");		
		}
		
	}


	public function create_supervisor()
	{

		if($this->session->userdata['admin_id'])
		{
			$name = $this->session->userdata['admin_name'];			
			//Redirect to dashbord

			$page_data['admin_name'] =  $name;
			$page_data['admin_id'] =  $this->session->userdata['admin_id'];
			$page_data['page_title'] = "admin Dashbord";
			$page_data['page'] = "admin/create_supervisor";
			$this->load->view('admin/layouts/dashboard', $page_data);
		}
		else
		{
			redirect("index.php/Admin/login");		
		}
	}

	public function add_supervisor()
	{
		if($this->input->post())
		{
			$this->load->model('main_model');
			$name = $this->input->post('name');
			$mobile = $this->input->post('mobile');
			$gender = $this->input->post('gender');
			$email = $this->input->post('email');
			$password = $this->input->post('password');

			$supervisor_data = array
			(
				"supervisor_name" =>$name,
				"supervisor_email" =>$email,
				"password" =>$password,
				"mobile" =>$mobile,
				"gender" =>$gender,
				"status" =>1,
			);
			$this->Admin_model->insert_supervisor($supervisor_data);
			redirect("index.php/Admin/all_supervisor");
		}
		else
		{
			redirect("inder.php/Admin/login");		
		}
	}
	public function all_supervisor()
	{	
		if($this->session->userdata['admin_id'])
		{
			$all_supervisor = $this->Admin_model->supervisors_list();
			$page_data['admin_name'] =  $this->session->userdata['admin_name'];
			$page_data['supervisors_data'] = $all_supervisor;
			$page_data['page_title'] = "Create User";
			$page_data['page'] = "admin/all_supervisors";
			$this->load->view('admin/layouts/dashboard', $page_data);
		}
		else
		{
			redirect("index.php/Admin/login");		
		}
		
	}

	public function supervisor_details($supervisor_id)
	{

		$supervisor_details = $this->Admin_model->supervisor_details($supervisor_id);

		// $supervisor_team_strength = $this->select->supervisor_team_strength($supervisor_id);
		// $page_data['supervisor_team_strength'] = $supervisor_team_strength;
		$page_data['admin_name'] =  $this->session->userdata['admin_name'];
		$page_data['supervisors_details'] = $supervisor_details;
		$page_data['page_title'] = "Create User";
		$page_data['page'] = "admin/supervisor_details";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}


	public function edit_supervisor($supervisor_id)
	{
		$supervisor_details = $this->Admin_model->supervisor_details($supervisor_id);
		$page_data['supervisor_details'] = $supervisor_details;
		$page_data['admin_name'] =  $this->session->userdata['admin_name'];
		$page_data['page'] = "admin/edit_supervisor";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}
	public function update_supervisor()
	{
		$name = $this->input->post('name');
		$mobile = $this->input->post('mobile');
		$gender = $this->input->post('gender');
		$email = $this->input->post('email');
		$password = $this->input->post('password');
		$supervisor_id = $this->input->post('supervisor_id');

		$supervisor_data = array(
			'supervisor_name' => $name,
			'supervisor_email' => $email,
			'password' => $password,
			'mobile' => $mobile,
			'gender' => $gender,
		);
		$this->Admin_model->edit_supervisor($supervisor_data, $supervisor_id);
		redirect("index.php/Admin/supervisor_details/".$supervisor_id);
	}



	public function deactivate_account($designation, $id)
	{
		

		if($designation == 'supervisor')
		{
			$supervisor_id = $id;

			$this->Admin_model->supervisor_deactivate($supervisor_id);
			redirect("index.php/Admin/supervisor_details/".$supervisor_id);
		}

		if($designation == 'agent')
		{

			$agent_id = $id;

			$this->Admin_model->agent_deactivate($agent_id);
			redirect("index.php/Admin/agent_details/".$agent_id);
		}


	}
	public function activate_account($designation, $id)
	{
		
		
		if($designation == 'supervisor')
		{
			$supervisor_id = $id;
			
			$this->Admin_model->supervisor_activate($supervisor_id);
			redirect("index.php/Admin/supervisor_details/".$supervisor_id);
		}

		if($designation == 'agent')
		{
			$agent_id = $id;
			
			$this->Admin_model->agent_activate($agent_id);
			redirect("index.php/Admin/agent_details/".$agent_id);
		}


	}


	// this fucntion will redirect to setting page
	public function setting()
	{
		$project_name = $this->session->userdata['project_name'];
		$name = $this->session->userdata['owner_name'];	
		$crm_id = $this->session->userdata['crm_id'];			

		$this->load->model('select');
		
		$page_data['input_menus'] = $this->select->input_menus($crm_id);
		$page_data['datatable_menus'] = $this->select->datatable_menus($crm_id);
		//Redirect to dashbord
		$page_data['project_name'] =  $project_name;
		$page_data['admin_name'] =  $name;
		$page_data['crm_id'] =  $crm_id;
		$page_data['page_title'] = "admin Dashbord";
		$page_data['page'] = "admin/setting";
		$this->load->view('admin/layouts/setting', $page_data);
	}

	public function view_profile()
	{
		$crm_id = $this->session->userdata['crm_id'];
		$page_data['profile_details'] = $this->Admin_model->profile_details($crm_id);

		// echo "<pre>";
		// print_r($page_data['profile_details']);
		// die();

		$project_name = $this->session->userdata['project_name'];
		$name = $this->session->userdata['owner_name'];	
		$crm_id = $this->session->userdata['crm_id'];			
		$page_data['project_name'] =  $project_name;
		$page_data['admin_name'] =  $name;
		$page_data['crm_id'] =  $crm_id;
		$page_data['page_title'] = "admin Dashbord";
		$page_data['page'] = "admin/view_profile";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}



	public function cust_info_data()
	{

		$page_data['cust_info_data'] = $this->Report_model->cust_info_data();

		$page_data['page_title'] = "Customer Data";
		$page_data['page'] = "admin/cust_info_data";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}
	public function todays_case()
	{

		$page_data['cust_info_data'] = $this->Report_model->todays_case();

		$page_data['page_title'] = "today cases";
		$page_data['page'] = "admin/cust_info_data";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}
	public function alcoholism_data()
	{

		$page_data['alcoholism_data'] = $this->Report_model->alcoholism_data();
		$page_data['page_title'] = "Alcoholism Data";
		$page_data['page'] = "admin/alcoholism_data";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}
	public function drug_abuse_data()
	{

		$page_data['drug_abuse_data'] = $this->Report_model->drug_abuse_data();
		
		$page_data['page_title'] = "Drugabuse data";
		$page_data['page'] = "admin/drug_abuse_data";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}
	public function alcohal_sale_data()
	{

		$page_data['alcohal_sale_data'] = $this->Report_model->alcohal_sale_data();
		$page_data['page_title'] = "Alcohal Sale Data";
		$page_data['page'] = "admin/alcohal_sale_data";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}
	public function drug_sale_data()
	{

		$page_data['drug_sale_data'] = $this->Report_model->drug_sale_data();

		$page_data['page_title'] = "Drug Sale Data";
		$page_data['page'] = "admin/drug_sale_data";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}
	public function other_data()
	{

		$page_data['other_data'] = $this->Report_model->other_data();

		$page_data['page_title'] = "Other Data";
		$page_data['page'] = "admin/other_data";
		$this->load->view('admin/layouts/dashboard', $page_data);
	}


	public function cust_info_exportexcel()
	{
		$start_date = $this->input->post('start_date');
		$end_date = $this->input->post('end_date');
		$start_date = date("d-M-y", strtotime($start_date));
		$end_date = date("d-M-y", strtotime($end_date));
		$fileName = 'customer Information.xlsx';  
		$cust_info_data = $this->Report_model->filtered_cust_info_data($start_date, $end_date);
		$spreadsheet = new Spreadsheet();
	        $sheet = $spreadsheet->getActiveSheet();
		$sheet->setCellValue('A1', 'Tickdasdet No');
        	$sheet->setCellValue('B1', 'Language');
	        $sheet->setCellValue('C1', 'Type of Caller');
	        $sheet->setCellValue('D1', 'Marital status');
		$sheet->setCellValue('E1', 'Occupation');
        	$sheet->setCellValue('F1', 'Age'); 
	        $sheet->setCellValue('G1', 'Gender'); 
        	$sheet->setCellValue('H1', 'Education status');  
	        $sheet->setCellValue('I1', 'Name'); 
	        $sheet->setCellValue('J1', 'Contact no'); 
        	$sheet->setCellValue('K1', 'State');  
	        $sheet->setCellValue('L1', 'District');  
	        $sheet->setCellValue('M1', 'Source of information');  
	        $sheet->setCellValue('N1', 'Type of call');  
	        $sheet->setCellValue('O1', 'Call related to');  
        	$sheet->setCellValue('P1', 'Call related to Others'); 
	        $sheet->setCellValue('Q1', 'Client Query 1');  
	        $sheet->setCellValue('R1', 'Reply Query 1');  
	        $sheet->setCellValue('S1', 'Client Query 2');  
	        $sheet->setCellValue('T1', 'Reply Query 2');  
	        $sheet->setCellValue('U1', 'Client Query 3');  
	        $sheet->setCellValue('V1', 'Reply Query 3');  
        $sheet->setCellValue('W1', 'Date'); 
        $sheet->setCellValue('X1', 'Time');  
        $sheet->setCellValue('Y1', 'Agent');  
        $rows = 2;
        foreach ($cust_info_data as $val){
            $sheet->setCellValue('A' . $rows, $val['callid']);
            $sheet->setCellValue('B' . $rows, $val['language']);
            $sheet->setCellValue('C' . $rows, $val['TypeofCaller']);
            $sheet->setCellValue('D' . $rows, $val['marital_status']);
	    $sheet->setCellValue('E' . $rows, $val['Occupation']);
            $sheet->setCellValue('F' . $rows, $val['age']);
            $sheet->setCellValue('G' . $rows, $val['gender']);
            $sheet->setCellValue('H' . $rows, $val['Educationstatus']);
            $sheet->setCellValue('I' . $rows, $val['name']);
            $sheet->setCellValue('J' . $rows, $val['contact_no']);
            $sheet->setCellValue('K' . $rows, $val['state']);
            $sheet->setCellValue('L' . $rows, $val['District']);
            $sheet->setCellValue('M' . $rows, $val['source_of_information']);
            $sheet->setCellValue('N' . $rows, $val['type_of_call']);
            $sheet->setCellValue('O' . $rows, $val['Call_related_to']);
            $sheet->setCellValue('P' . $rows, $val['Call_related_toOthers']);
            $sheet->setCellValue('Q' . $rows, $val['ClientQuery_1']);
            $sheet->setCellValue('R' . $rows, $val['replyQuery_1']);
            $sheet->setCellValue('S' . $rows, $val['ClientQuery_2']);
            $sheet->setCellValue('T' . $rows, $val['replyQuery_2']);
            $sheet->setCellValue('U' . $rows, $val['ClientQuery_3']);
            $sheet->setCellValue('V' . $rows, $val['replyQuery_3']);
            $sheet->setCellValue('W' . $rows, $val['Date']);
            $sheet->setCellValue('X' . $rows, $val['time']);
            $sheet->setCellValue('Y' . $rows, $val['Agent']);
            $rows++;
        } 
        $writer = new Xlsx($spreadsheet);
		$writer->save("Antidrug ".$fileName);
		header("Content-Type: application/vnd.ms-excel");
        redirect(base_url()."Antidrug ".$fileName);  
	}
	public function other_exportexcel()
	{
		$start_date = $this->input->post('start_date');
		$end_date = $this->input->post('end_date');
		$start_date = date("d-M-y", strtotime($start_date));
		$end_date = date("d-M-y", strtotime($end_date));

		$fileName = 'Proxy Data.xlsx'; 
	
 
		$cust_info_data = $this->Report_model->filtered_other_data($start_date, $end_date);

		$spreadsheet = new Spreadsheet();

        	$sheet = $spreadsheet->getActiveSheet();

		$sheet->setCellValue('A1', 'Callid');
	        $sheet->setCellValue('B1', 'Language');
	        $sheet->setCellValue('C1', 'TypeofCaller');
        	$sheet->setCellValue('D1', 'marital_status');
		$sheet->setCellValue('E1', 'Occupation');
	        $sheet->setCellValue('F1', 'Age'); 
        	$sheet->setCellValue('G1', 'Gender'); 
	        $sheet->setCellValue('H1', 'Educationstatus');  
        	$sheet->setCellValue('I1', 'Name'); 
	        $sheet->setCellValue('J1', 'contact_no'); 
       		$sheet->setCellValue('K1', 'state');  
        $sheet->setCellValue('L1', 'district');  
        $sheet->setCellValue('M1', 'address');  
        $sheet->setCellValue('N1', 'customer_query');  
        $sheet->setCellValue('O1', 'Reply Given By Counseller(Level-1)');  
        $sheet->setCellValue('P1', 'customer_query'); 
        $sheet->setCellValue('Q1', 'R eply Given By Counseller(Level-1)');  
        $sheet->setCellValue('R1', 'customer_query');  
        $sheet->setCellValue('S1', 'Reply Given By Counseller(Level-1)');  
        $sheet->setCellValue('T1', 'customer_query');  
        $sheet->setCellValue('U1', 'Reply Given By Counseller(Level-1)');  
        $sheet->setCellValue('V1', 'source_of_information');  
        $sheet->setCellValue('W1', 'Call_related_to'); 
        $sheet->setCellValue('X1', 'type_of_call');  
        $sheet->setCellValue('Y1', 'Date'); 
        $sheet->setCellValue('Z1', 'time');  
        $sheet->setCellValue('AA1', 'Agent');
        $rows = 2;

        foreach ($cust_info_data as $val){
        	if($val['Call_related_to'] == 'Proxy/Prank Call')
			{
	            $sheet->setCellValue('A' . $rows, $val['callid']);
	            $sheet->setCellValue('B' . $rows, $val['language']);
	            $sheet->setCellValue('C' . $rows, $val['TypeofCaller']);
	            $sheet->setCellValue('D' . $rows, $val['marital_status']);
		    $sheet->setCellValue('E' . $rows, $val['Occupation']);
	            $sheet->setCellValue('F' . $rows, $val['age']);
	            $sheet->setCellValue('G' . $rows, $val['gender']);
	            $sheet->setCellValue('H' . $rows, $val['Educationstatus']);
	            $sheet->setCellValue('I' . $rows, $val['name']);
	            $sheet->setCellValue('J' . $rows, $val['contact_no']);
	            $sheet->setCellValue('K' . $rows, $val['state']);
	            $sheet->setCellValue('L' . $rows, $val['district']);
	            $sheet->setCellValue('M' . $rows, $val['address']);
	            $sheet->setCellValue('N' . $rows, $val['']);
	            $sheet->setCellValue('O' . $rows, $val['']);
	            $sheet->setCellValue('P' . $rows, $val['ClientQuery_1']);
	            $sheet->setCellValue('Q' . $rows, $val['replyQuery_1']);
	            $sheet->setCellValue('R' . $rows, $val['ClientQuery_2']);
	            $sheet->setCellValue('S' . $rows, $val['replyQuery_2']);
	            $sheet->setCellValue('T' . $rows, $val['ClientQuery_3']);
	            $sheet->setCellValue('U' . $rows, $val['replyQuery_3']);
	            $sheet->setCellValue('V' . $rows, $val['source_of_information']);
	            $sheet->setCellValue('W' . $rows, $val['Call_related_to']);
	            $sheet->setCellValue('X' . $rows, $val['type_of_call']);
	            $sheet->setCellValue('Y' . $rows, $val['Date']);
	            $sheet->setCellValue('Z' . $rows, $val['time']);
	            $sheet->setCellValue('AA' . $rows, $val['Agent']);

	            $rows++;
	        }
        	else
        	{
        		
        	}
        } 

        $writer = new Xlsx($spreadsheet);
		$writer->save("Antidrug ".$fileName);
		header("Content-Type: application/vnd.ms-excel");
        redirect(base_url()."Antidrug ".$fileName);  
	}

	public function alcoholism_data_exportexcel()
	{
		$start_date = $this->input->post('start_date');
		$end_date = $this->input->post('end_date');
		$start_date = date("d-M-y", strtotime($start_date));
		$end_date = date("d-M-y", strtotime($end_date));
		$fileName = 'Alcolism Data.xlsx';  
		$alcoholism_data = $this->Report_model->filtered_alcoholism_data($start_date, $end_date);
		// echo "<pre>";
		// print_r($alcoholism_data);
		// die();
		$spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();   
		$sheet->setCellValue('A1', 'Callid');
		$sheet->setCellValue('B1', 'TypeofCaller');
		$sheet->setCellValue('C1', 'marital_status');
		$sheet->setCellValue('D1', 'Occupation');
		$sheet->setCellValue('E1', 'age');
		$sheet->setCellValue('F1', 'gender');
		$sheet->setCellValue('G1', 'Educationstatus');
		$sheet->setCellValue('H1', 'name');
		$sheet->setCellValue('I1', 'contact_no');
		$sheet->setCellValue('J1', 'state');
		$sheet->setCellValue('K1', 'district');
		$sheet->setCellValue('L1', 'address');
		$sheet->setCellValue('M1', 'source_of_information');
		$sheet->setCellValue('N1', 'Call_related_to');
		$sheet->setCellValue('O1', 'type_of_call');
		$sheet->setCellValue('P1', 'information');
        $sheet->setCellValue('Q1', 'ClientLIvingStatus');
		$sheet->setCellValue('R1', 'Date');
		$sheet->setCellValue('S1', 'time');
		$sheet->setCellValue('T1', 'Agent');
		$sheet->setCellValue('U1', 'IfAlcoholism:Mode/Frequency of consumption');
		$sheet->setCellValue('V1', 'customer_query');
		$sheet->setCellValue('W1', 'Reply Given By Counseller(Level-1)');
        $sheet->setCellValue('X1', 'customer_query');
        $sheet->setCellValue('Y1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('Z1', 'customer_query');
        $sheet->setCellValue('AA1', 'Reply Given By Counseller(Level-1)');  
        $sheet->setCellValue('AB1', 'customer_query');  
        $sheet->setCellValue('AC1', 'Reply Given By Counseller(Level-1)');  
        $sheet->setCellValue('AD1', 'IfAlcoholism:Any history of alcoholic in the Family');  
        $sheet->setCellValue('AE1', 'Treatmentearlier1');  
        $sheet->setCellValue('AF1', 'IfAlcoholism:If Yes, Relation with Alcoholic family member');  
        $sheet->setCellValue('AG1', 'IfAlcoholism:Problem Identified by the Counsellor');  
        $sheet->setCellValue('AH1', 'IfAlcoTreatmentYes'); 
        $rows = 2;
        // echo "<pre>";print_r($alcoholism_data);die();
        foreach ($alcoholism_data as $val){

        	if($val['Call_related_to'] == 'Alcoholism')
			{
				$sheet->setCellValue('A' . $rows, $val['callid']);
	            $sheet->setCellValue('B' . $rows, $val['TypeofCaller']);
	            $sheet->setCellValue('C' . $rows, $val['marital_status']);
	            $sheet->setCellValue('D' . $rows, $val['Occupation']);
				$sheet->setCellValue('E' . $rows, $val['age']);
	            $sheet->setCellValue('F' . $rows, $val['gender']);
	            $sheet->setCellValue('G' . $rows, $val['Educationstatus']);
	            $sheet->setCellValue('H' . $rows, $val['name']);
	            $sheet->setCellValue('I' . $rows, $val['contact_no']);
	            $sheet->setCellValue('J' . $rows, $val['state']);
	            $sheet->setCellValue('K' . $rows, $val['district']);
	            $sheet->setCellValue('L' . $rows, $val['address']);
	            $sheet->setCellValue('M' . $rows, $val['source_of_information']);
	            $sheet->setCellValue('N' . $rows, $val['Call_related_to']);
	            $sheet->setCellValue('O' . $rows, $val['type_of_call']);
	            $sheet->setCellValue('P' . $rows, $val['type_of_information_alcohalism']);
	            $sheet->setCellValue('Q' . $rows, $val['ClientLIvingStatus']);
	            $sheet->setCellValue('R' . $rows, $val['Date']);
	            $sheet->setCellValue('S' . $rows, $val['time']);
	            $sheet->setCellValue('T' . $rows, $val['Agent']);
	            $sheet->setCellValue('U' . $rows, $val['Frequency_of_consumption1']);
	            $sheet->setCellValue('V' . $rows, $val['ClientQuery']);
	            $sheet->setCellValue('W' . $rows, $val['replyQuery']);
	            $sheet->setCellValue('X' . $rows, $val['ClientQuery_1']);
	            $sheet->setCellValue('Y' . $rows, $val['replyQuery_1']);
	            $sheet->setCellValue('Z' . $rows, $val['ClientQuery_2']);
	            $sheet->setCellValue('AA' . $rows, $val['replyQuery_2']);
	            $sheet->setCellValue('AB' . $rows, $val['ClientQuery_3']);
	            $sheet->setCellValue('AC' . $rows, $val['replyQuery_3']);
	            $sheet->setCellValue('AD' . $rows, $val['Historyalcoholic']);
	            $sheet->setCellValue('AE' . $rows, $val['Treatmentearlier1']);
	            $sheet->setCellValue('AF' . $rows, $val['ResOnAlcho']);
	            $sheet->setCellValue('AG' . $rows, $val['ifAlcoProblem_Identified']);
	            $sheet->setCellValue('AH' . $rows, $val['IfAlcoTreatmentYes']);

	            $rows++;
        	}
        	else
        	{
        		
        	}
        } 

        $writer = new Xlsx($spreadsheet);
        
		$writer->save("Antidrug ".$fileName);
		
		header("Content-Type: application/vnd.ms-excel");
        redirect(base_url()."Antidrug ".$fileName);  

                 
	}


	public function drugabuse_data_exportexcel()
	{
		$start_date = $this->input->post('start_date');
		$end_date = $this->input->post('end_date');
		$start_date = date("d-M-y", strtotime($start_date));
		$end_date = date("d-M-y", strtotime($end_date));
		$fileName = 'Drugabuse Data.xlsx';  
		$drug_abuse_data = $this->Report_model->filtered_drug_abuse_data($start_date, $end_date);
		$spreadsheet = new Spreadsheet();
 	     	$sheet = $spreadsheet->getActiveSheet();   
		$sheet->setCellValue('A1', 'Callid');
		$sheet->setCellValue('B1', 'language');
		$sheet->setCellValue('C1', 'TypeofCaller');
		$sheet->setCellValue('D1', 'marital_status');
		$sheet->setCellValue('E1', 'Occupation');
		$sheet->setCellValue('F1', 'age');
		$sheet->setCellValue('G1', 'gender');
		$sheet->setCellValue('H1', 'Educationstatus');
		$sheet->setCellValue('I1', 'name');
		$sheet->setCellValue('J1', 'contact_no');
		$sheet->setCellValue('K1', 'state');
		$sheet->setCellValue('L1', 'district');
		$sheet->setCellValue('M1', 'address');
		$sheet->setCellValue('N1', 'customer_query');
		$sheet->setCellValue('O1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('P1', 'customer_query');
        	$sheet->setCellValue('Q1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('R1', 'customer_query');
		$sheet->setCellValue('S1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('T1', 'customer_query');
		$sheet->setCellValue('U1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('V1', 'source_of_information');
		$sheet->setCellValue('W1', 'Call_related_to');
        	$sheet->setCellValue('X1', 'type_of_call');
        	$sheet->setCellValue('Y1', 'Date');
		$sheet->setCellValue('Z1', 'time');
        	$sheet->setCellValue('AA1', 'Agent');  
        	$sheet->setCellValue('AB1', 'information');  
        	$sheet->setCellValue('AC1', 'grievance');  
        	$sheet->setCellValue('AD1', 'type of drug');  
        	$sheet->setCellValue('AE1', 'Any history of drug abuse in the family');  
        	$sheet->setCellValue('AF1', 'Whether an injecting drug used');  
        	$sheet->setCellValue('AG1', 'Mode/Frequency of consumption');  
        	$sheet->setCellValue('AH1', 'Whether treatment taken earlier'); 
        	$sheet->setCellValue('AI1', 'Whether treatment taken earlier:If yes :Where'); 
        	$sheet->setCellValue('AJ1', 'Problem Identified by the Counsellor'); 
        	$sheet->setCellValue('AK1', 'information_Distict'); 
        	$sheet->setCellValue('AL1', 'information_state'); 
        	$sheet->setCellValue('AM1', 'Information_ngo'); 
        	$sheet->setCellValue('AN1', 'TypOfDrugConsum'); 
        	$sheet->setCellValue('AO1', 'NGO_Name'); 
        	$sheet->setCellValue('AP1', 'Center_Location'); 
        	$sheet->setCellValue('AQ1', 'Contact_Number'); 
        	$sheet->setCellValue('AR1', 'Email_ID'); 
	
        $rows = 2;
        foreach ($drug_abuse_data as $val){

        	if($val['Call_related_to'] == 'Drug Abuse')
		{
		    $sheet->setCellValue('A' . $rows, $val['callid']);
	            $sheet->setCellValue('B' . $rows, $val['language']);
	            $sheet->setCellValue('C' . $rows, $val['TypeofCaller']);
	            $sheet->setCellValue('D' . $rows, $val['marital_status']);
				$sheet->setCellValue('E' . $rows, $val['Occupation']);
	            $sheet->setCellValue('F' . $rows, $val['age']);
	            $sheet->setCellValue('G' . $rows, $val['gender']);
	            $sheet->setCellValue('H' . $rows, $val['Educationstatus']);
	            $sheet->setCellValue('I' . $rows, $val['name']);
	            $sheet->setCellValue('J' . $rows, $val['contact_no']);
	            $sheet->setCellValue('K' . $rows, $val['state']);
	            $sheet->setCellValue('L' . $rows, $val['district']);
	            $sheet->setCellValue('M' . $rows, $val['address']);
	            $sheet->setCellValue('N' . $rows, $val['customer_query']);
	            $sheet->setCellValue('O' . $rows, $val['rply_given_by_counseller']);
	            $sheet->setCellValue('P' . $rows, $val['ClientQuery_1']);
	            $sheet->setCellValue('Q' . $rows, $val['replyQuery_1']);
	            $sheet->setCellValue('R' . $rows, $val['ClientQuery_2']);
	            $sheet->setCellValue('S' . $rows, $val['replyQuery_2']);
	            $sheet->setCellValue('T' . $rows, $val['ClientQuery_3']);
	            $sheet->setCellValue('U' . $rows, $val['replyQuery_3']);
	            $sheet->setCellValue('V' . $rows, $val['source_of_information']);
	            $sheet->setCellValue('W' . $rows, $val['Call_related_to']);
	            $sheet->setCellValue('X' . $rows, $val['type_of_call']);
	            $sheet->setCellValue('Y' . $rows, $val['Date']);
	            $sheet->setCellValue('Z' . $rows, $val['time']);
	            $sheet->setCellValue('AA' . $rows, $val['Agent']);
	            $sheet->setCellValue('AB' . $rows, $val['information']);
	            $sheet->setCellValue('AC' . $rows, $val['grievance']);
	            $sheet->setCellValue('AD' . $rows, $val['TypOfDrugConsum']);
	            $sheet->setCellValue('AE' . $rows, $val['history_of_drug_abuse']);
	            $sheet->setCellValue('AF' . $rows, $val['injecting_drug']);
	            $sheet->setCellValue('AG' . $rows, $val['Frequency_of_consumption']);
	            $sheet->setCellValue('AH' . $rows, $val['Treatmentearlier']);
	            $sheet->setCellValue('AI' . $rows, $val['ifdrugTrestmentYes']);
	            $sheet->setCellValue('AJ' . $rows, $val['ifDrugProblem_Identified']);
	            $sheet->setCellValue('AK' . $rows, $val['information_Distict']);
	            $sheet->setCellValue('AL' . $rows, $val['information_state']);
	            $sheet->setCellValue('AM' . $rows, $val['Information_ngo']);
	            $sheet->setCellValue('AN' . $rows, $val['TypOfDrugConsum']);
	            $sheet->setCellValue('AO' . $rows, $val['NGO_Name']);
	            $sheet->setCellValue('AP' . $rows, $val['Center_Location']);
	            $sheet->setCellValue('AQ' . $rows, $val['Contact_Number']);
	            $sheet->setCellValue('AR' . $rows, $val['Email_ID']);


	            $rows++;
        	}
        	else
        	{
        		
        	}
        } 

        $writer = new Xlsx($spreadsheet);
        
		$writer->save("Antidrug ".$fileName);
		
		header("Content-Type: application/vnd.ms-excel");
        redirect(base_url()."Antidrug ".$fileName);  

                 
	}


	public function drugsale_data_exportexcel()
	{
		$start_date = $this->input->post('start_date');
		$end_date = $this->input->post('end_date');
		$start_date = date("d-M-y", strtotime($start_date));
		$end_date = date("d-M-y", strtotime($end_date));
		$fileName = 'Drug Sale Data.xlsx';  
		$drug_sale_data = $this->Report_model->filtered_drug_sale_data($start_date, $end_date);
		// echo "<pre>";
		// print_r($alcoholism_data);
		// die();
		$spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();   
		$sheet->setCellValue('A1', 'callid');
		$sheet->setCellValue('B1', 'language');
		$sheet->setCellValue('C1', 'TypeofCaller');
		$sheet->setCellValue('D1', 'marital_status');
		$sheet->setCellValue('E1', 'Occupation');
		$sheet->setCellValue('F1', 'age');
		$sheet->setCellValue('G1', 'gender');
		$sheet->setCellValue('H1', 'Educationstatus');
		$sheet->setCellValue('I1', 'name');
		$sheet->setCellValue('J1', 'contact_no');
		$sheet->setCellValue('K1', 'state');
		$sheet->setCellValue('L1', 'district');
		$sheet->setCellValue('M1', 'address');
		$sheet->setCellValue('N1', 'customer_query');
		$sheet->setCellValue('O1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('P1', 'customer_query');
        $sheet->setCellValue('Q1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('R1', 'customer_query');
		$sheet->setCellValue('S1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('T1', 'customer_query');
		$sheet->setCellValue('U1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('V1', 'source_of_information');
		$sheet->setCellValue('W1', 'Call_related_to');
        $sheet->setCellValue('X1', 'Date');
        $sheet->setCellValue('Y1', 'time');
		$sheet->setCellValue('Z1', 'Agent');
        $sheet->setCellValue('AA1', 'A)Kind of drug in sale');  
        $sheet->setCellValue('AB1', 'B)Address /area /location where drug is being sold');  
        $sheet->setCellValue('AC1', 'C)Specific time of sale');  
        $sheet->setCellValue('AD1', 'D)Name of the suspected person');  
        $sheet->setCellValue('AE1', 'E)Address of suspected person');  
        $sheet->setCellValue('AF1', 'F)Method being used by person for sale');  
        $sheet->setCellValue('AG1', 'G)Current profession of suspected person');  
        $sheet->setCellValue('AH1', 'H)Financial condition of suspected personn'); 
        $sheet->setCellValue('AI1', 'I)Does the suspected person owns any vehicle?'); 
        $sheet->setCellValue('AJ1', 'J)Whether suspected person is being jailed'); 
        $sheet->setCellValue('AK1', 'K)Any Associate of the suspected person'); 
        $sheet->setCellValue('AL1', 'If yes for K) above: Details about the associate such as Name,address,contact details etc.'); 
        $sheet->setCellValue('AM1', 'L) Any Additional Information'); 
        $sheet->setCellValue('AN1', 'saleCmpltYes'); 
        $sheet->setCellValue('AO1', 'saleComplaint'); 
        $sheet->setCellValue('AP1', 'saleSource'); 
        $sheet->setCellValue('AQ1', 'saleMobile'); 
        $sheet->setCellValue('AR1', 'IfsaleName'); 
        $sheet->setCellValue('AS1', 'saleYesNo'); 
       
        $rows = 2;
        // echo "<pre>";print_r($alcoholism_data);die();
        foreach ($drug_sale_data as $val){

        	if($val['Call_related_to'] == 'Drug Sale')
		{
		$sheet->setCellValue('A' . $rows, $val['callid']);
	            $sheet->setCellValue('B' . $rows, $val['language']);
	            $sheet->setCellValue('C' . $rows, $val['TypeofCaller']);
	            $sheet->setCellValue('D' . $rows, $val['marital_status']);
				$sheet->setCellValue('E' . $rows, $val['Occupation']);
	            $sheet->setCellValue('F' . $rows, $val['age']);
	            $sheet->setCellValue('G' . $rows, $val['gender']);
	            $sheet->setCellValue('H' . $rows, $val['Educationstatus']);
	            $sheet->setCellValue('I' . $rows, $val['name']);
	            $sheet->setCellValue('J' . $rows, $val['contact_no']);
	            $sheet->setCellValue('K' . $rows, $val['state']);
	            $sheet->setCellValue('L' . $rows, $val['district']);
	            $sheet->setCellValue('M' . $rows, $val['address']);
	            $sheet->setCellValue('N' . $rows, $val['customer_query']);
	            $sheet->setCellValue('O' . $rows, $val['rply_given_by_counseller']);
	            $sheet->setCellValue('P' . $rows, $val['ClientQuery_1']);
	            $sheet->setCellValue('Q' . $rows, $val['replyQuery_1']);
	            $sheet->setCellValue('R' . $rows, $val['ClientQuery_2']);
	            $sheet->setCellValue('S' . $rows, $val['replyQuery_2']);
	            $sheet->setCellValue('T' . $rows, $val['ClientQuery_3']);
	            $sheet->setCellValue('U' . $rows, $val['replyQuery_3']);
	            $sheet->setCellValue('V' . $rows, $val['source_of_information']);
	            $sheet->setCellValue('W' . $rows, $val['Call_related_to']);
	            $sheet->setCellValue('X' . $rows, $val['Date']);
	            $sheet->setCellValue('Y' . $rows, $val['time']);
	            $sheet->setCellValue('Z' . $rows, $val['Agent']);
	            $sheet->setCellValue('AA' . $rows, $val['CMpailnA']);
	            $sheet->setCellValue('AB' . $rows, $val['CMpailnB']);
	            $sheet->setCellValue('AC' . $rows, $val['CMpailnC']);
	            $sheet->setCellValue('AD' . $rows, $val['CMpailnD']);
	            $sheet->setCellValue('AE' . $rows, $val['CMpailnE']);
	            $sheet->setCellValue('AF' . $rows, $val['CMpailnF']);
	            $sheet->setCellValue('AG' . $rows, $val['CMpailnG']);
	            $sheet->setCellValue('AH' . $rows, $val['CMpailnH']);
	            $sheet->setCellValue('AI' . $rows, $val['CMpailnI']);
	            $sheet->setCellValue('AJ' . $rows, $val['CMpailnJ']);
	            $sheet->setCellValue('AK' . $rows, $val['CMpailnK']);
	            $sheet->setCellValue('AL' . $rows, $val['CMpailnKyes']);
	            $sheet->setCellValue('AM' . $rows, $val['CMpailnL']);
	            $sheet->setCellValue('AN' . $rows, $val['saleCmpltYes']);
	            $sheet->setCellValue('AO' . $rows, $val['saleYesNo']);
	            $sheet->setCellValue('AP' . $rows, $val['saleSource']);
	            $sheet->setCellValue('AQ' . $rows, $val['saleMobile']);
	            $sheet->setCellValue('AR' . $rows, $val['IfsaleName']);
	            $sheet->setCellValue('AS' . $rows, $val['saleComplaint']);


	            $rows++;
        	}
        	else
        	{
        		
        	}
	        } 
	        $writer = new Xlsx($spreadsheet);
		$writer->save("Antidrug ".$fileName);
		header("Content-Type: application/vnd.ms-excel");
        	redirect(base_url()."Antidrug ".$fileName);  
	}



	public function alcohal_sale_data_exportexcel()
	{
		$start_date = $this->input->post('start_date');
		$end_date = $this->input->post('end_date');
		$start_date = date("d-M-y", strtotime($start_date));
		$end_date = date("d-M-y", strtotime($end_date));
		$fileName = 'Alcohal Sale Data.xlsx';  
		$alcohal_sale_data = $this->Report_model->filtered_alcohal_sale_data($start_date, $end_date);
	
		$spreadsheet = new Spreadsheet();
	        $sheet = $spreadsheet->getActiveSheet();   
		$sheet->setCellValue('A1', 'callid');
		$sheet->setCellValue('B1', 'language');
		$sheet->setCellValue('C1', 'TypeofCaller');
		$sheet->setCellValue('D1', 'marital_status');
		$sheet->setCellValue('E1', 'Occupation');
		$sheet->setCellValue('F1', 'age');
		$sheet->setCellValue('G1', 'gender');
		$sheet->setCellValue('H1', 'Educationstatus');
		$sheet->setCellValue('I1', 'name');
		$sheet->setCellValue('J1', 'contact_no');
		$sheet->setCellValue('K1', 'state');
		$sheet->setCellValue('L1', 'district');
		$sheet->setCellValue('M1', 'address');
		$sheet->setCellValue('N1', 'customer_query');
		$sheet->setCellValue('O1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('P1', 'customer_query');
        	$sheet->setCellValue('Q1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('R1', 'customer_query');
		$sheet->setCellValue('S1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('T1', 'customer_query');
		$sheet->setCellValue('U1', 'Reply Given By Counseller(Level-1)');
		$sheet->setCellValue('V1', 'source_of_information');
		$sheet->setCellValue('W1', 'Call_related_to');
	        $sheet->setCellValue('X1', 'Date');
	        $sheet->setCellValue('Y1', 'time');
		$sheet->setCellValue('Z1', 'Agent');
        	$sheet->setCellValue('AA1', 'AlcosaleYes');  
	        $sheet->setCellValue('AB1', 'NameAlco');  
	        $sheet->setCellValue('AC1', 'ContactAlco');  
        	$sheet->setCellValue('AD1', 'CmlnttAlco');  
	        $sheet->setCellValue('AE1', 'officialAlco');  
        	$sheet->setCellValue('AF1', 'officialYesAlco');  
	        $sheet->setCellValue('AG1', 'A)Kind of Alcohol in sale');  
        	$sheet->setCellValue('AH1', 'B)Address /area /location where drug is being sold'); 
	        $sheet->setCellValue('AI1', 'C)Specific time of sale'); 
        	$sheet->setCellValue('AJ1', 'D)Name of the suspected person'); 
	        $sheet->setCellValue('AK1', 'E)Address of suspected person'); 
	        $sheet->setCellValue('AL1', 'F)Method being used by person for sale'); 
	        $sheet->setCellValue('AM1', 'G)Current profession of suspected person'); 
	        $sheet->setCellValue('AN1', 'H)Financial condition of suspected personn'); 
	        $sheet->setCellValue('AO1', 'I)Does the suspected person owns any vehicle?'); 
	        $sheet->setCellValue('AP1', 'J)Whether suspected person is being jailed'); 
        	$sheet->setCellValue('AQ1', 'K)Any Associate of the suspected person'); 
	        $sheet->setCellValue('AR1', 'If yes for K) above: Details about the associate such as Name,address,contact details etc.'); 
	        $sheet->setCellValue('AS1', 'L) Any Additional Information'); 
	        $rows = 2;
        	// echo "<pre>";print_r($alcoholism_data);die();
	        foreach ($alcohal_sale_data as $val){

        	if($val['Call_related_to'] == 'Alcohol Sale')
			{
				$sheet->setCellValue('A' . $rows, $val['callid']);
	            $sheet->setCellValue('B' . $rows, $val['language']);
	            $sheet->setCellValue('C' . $rows, $val['TypeofCaller']);
	            $sheet->setCellValue('D' . $rows, $val['marital_status']);
				$sheet->setCellValue('E' . $rows, $val['Occupation']);
	            $sheet->setCellValue('F' . $rows, $val['age']);
	            $sheet->setCellValue('G' . $rows, $val['gender']);
	            $sheet->setCellValue('H' . $rows, $val['Educationstatus']);
	            $sheet->setCellValue('I' . $rows, $val['name']);
	            $sheet->setCellValue('J' . $rows, $val['contact_no']);
	            $sheet->setCellValue('K' . $rows, $val['state']);
	            $sheet->setCellValue('L' . $rows, $val['district']);
	            $sheet->setCellValue('M' . $rows, $val['address']);
	            $sheet->setCellValue('N' . $rows, $val['customer_query']);
	            $sheet->setCellValue('O' . $rows, $val['rply_given_by_counseller']);
	            $sheet->setCellValue('P' . $rows, $val['ClientQuery_1']);
	            $sheet->setCellValue('Q' . $rows, $val['replyQuery_1']);
	            $sheet->setCellValue('R' . $rows, $val['ClientQuery_2']);
	            $sheet->setCellValue('S' . $rows, $val['replyQuery_2']);
	            $sheet->setCellValue('T' . $rows, $val['ClientQuery_3']);
	            $sheet->setCellValue('U' . $rows, $val['replyQuery_3']);
	            $sheet->setCellValue('V' . $rows, $val['source_of_information']);
	            $sheet->setCellValue('W' . $rows, $val['Call_related_to']);
	            $sheet->setCellValue('X' . $rows, $val['Date']);
	            $sheet->setCellValue('Y' . $rows, $val['time']);
	            $sheet->setCellValue('Z' . $rows, $val['Agent']);
	            $sheet->setCellValue('AA' . $rows, $val['AlcosaleYes']);
	            $sheet->setCellValue('AB' . $rows, $val['NameAlco']);
	            $sheet->setCellValue('AC' . $rows, $val['ContactAlco']);
	            $sheet->setCellValue('AD' . $rows, $val['CmlnttAlco']);
	            $sheet->setCellValue('AE' . $rows, $val['officialAlco']);
	            $sheet->setCellValue('AF' . $rows, $val['officialYesAlco']);
	            $sheet->setCellValue('AG' . $rows, $val['CMpailnAAlco']);
	            $sheet->setCellValue('AH' . $rows, $val['CMpailnBAlco']);
	            $sheet->setCellValue('AI' . $rows, $val['CMpailnCAlco']);
	            $sheet->setCellValue('AJ' . $rows, $val['CMpailnDAlco']);
	            $sheet->setCellValue('AK' . $rows, $val['CMpailnEAlco']);
	            $sheet->setCellValue('AL' . $rows, $val['CMpailnFAlco']);
	            $sheet->setCellValue('AM' . $rows, $val['CMpailnGAlco']);
	            $sheet->setCellValue('AN' . $rows, $val['CMpailnHAlco']);
	            $sheet->setCellValue('AO' . $rows, $val['CMpailnIAlco']);
	            $sheet->setCellValue('AP' . $rows, $val['CMpailnJAlco']);
	            $sheet->setCellValue('AQ' . $rows, $val['CMpailnKAlco']);
	            $sheet->setCellValue('AR' . $rows, $val['CMpailnKLAlco']);
	            $sheet->setCellValue('AS' . $rows, $val['CMpailnKyesAlco']);


	            $rows++;
        	}
        	else
        	{
        		
        	}
        } 

        $writer = new Xlsx($spreadsheet);
        
		$writer->save("Antidrug ".$fileName);
		
		header("Content-Type: application/vnd.ms-excel");
        redirect(base_url()."Antidrug ".$fileName);  

                 
	}


	public function logout()
	{
		$this->session->sess_destroy();
		redirect('index.php/Admin');
		
	}
	


}
?>