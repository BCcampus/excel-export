<?php

namespace BCcampus\Excel\Users;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 *
 * @return string
 * @throws \PhpOffice\PhpSpreadsheet\Exception
 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
 */
function excel_export_users() {
	// check if User data is being requested and nonce is valid
	if ( ! empty( $_POST ) && isset( $_POST['users_export'] ) && check_admin_referer( 'export_button_users', 'submit_export_users' ) ) {

		// Create a new spreadsheet
		$spreadsheet = new Spreadsheet();

		// Args for the user query
		$args = [
			'order'   => 'ASC',
			'orderby' => 'display_name',
			'fields'  => 'all',
		];

		// User Query
		$wp_users   = get_users( $args );
		$cell_count = 1;

		// BuddyPress user data placeholders
		$bp_field_ids   = [];
		$bp_field_names = [];
		$bp_field_data  = [];

		// Get BuddyPress profile field ID's and names
		if ( function_exists( 'bp_is_active' ) ) {

			$profile_groups = \BP_XProfile_Group::get(
				[
					'fetch_fields' => true,
				]
			);

			if ( ! empty( $profile_groups ) ) {
				foreach ( $profile_groups as $profile_group ) {
					if ( ! empty( $profile_group->fields ) ) {
						foreach ( $profile_group->fields as $field ) {
							$bp_field_names[] = $field->name;
							$bp_field_ids[]   = $field->id;
						}
					}
				}
			}
		}

		// don't include personally identifiable information in export by default
		( isset( $_POST['users_export'] ) ) ? $consent = $_POST['consent'] : $consent = '0';
		$alphabet                                      = range( 'A', 'Z' );
		$user_data                                     = [
			'ID'              => 'User ID',
			'user_login'      => 'Username',
			'display_name'    => 'Display Name',
			'first_name'      => 'First Name',
			'last_name'       => 'Last Name',
			'user_email'      => 'Email',
			'user_url'        => 'URL',
			'user_registered' => 'Registration Date',
			'roles'           => 'Roles',
			'user_level'      => 'User Level',
			'user_status'     => 'User Status',
		];

		// let developers hook in
		apply_filters( 'excel_export_user_data', $user_data );

		// no metadata by default
		$user_metadata = [];

		// let developers hook in
		apply_filters( 'excel_export_user_metadata', $user_metadata );

		// combine user and user_meta arrays
		$combined = array_merge( $user_data, $user_metadata );

		$num_columns = count( $combined );

		// get keys for only what we need
		$alpha_keys = array_splice( $alphabet, 0, $num_columns );

		// create cell headers from the alpha keys and filtered values
		$cell_headers = array_combine( $alpha_keys, array_values( $combined ) );

		// set cell headers
		foreach ( $cell_headers as $k => $v ) {
			$spreadsheet->setActiveSheetIndex( 0 )
						->SetCellValue( $k . $cell_count, $v );
		}

		/**
		 * Get User data for each user
		 */
		foreach ( $wp_users as $user ) {
			$cell_count ++;

			// create dynamic array based on what we might expect to be held in the user database
			$user_fields = array_combine( $alpha_keys, array_keys( $user_data ) );

			$user_content = get_from_users_table( $user->ID, $user_fields, $consent );

			if ( function_exists( 'bp_is_active' ) ) {
				// Get the BP data for this user
				$bp_get_data = \BP_XProfile_ProfileData::get_data_for_user( $user->ID, $bp_field_ids );

				// Get the value of BP fields for this user
				foreach ( $bp_get_data as $bp_field_value ) {
					$bp_field_data [] = $bp_field_value->value;
				}
			}

			// set csv basic user data
			foreach ( $user_content as $k => $v ) {
				$spreadsheet->setActiveSheetIndex( 0 )
							->SetCellValue( $k . $cell_count, $v );
			}

			// Get all the user meta into an array, run array_map to take only the first index of each result
			$user_meta = array_map(
				function ( $a ) {
					return $a[0];
				}, get_user_meta( $user->ID )
			);

			// remove session tokens value as a preventative security measure
			if ( isset( $user_meta['session_tokens'] ) ) {
				unset( $user_meta['session_tokens'] );
			}
		}
		// get column labels, user_id 1 as a placeholder for all fields
		$user_meta = get_user_meta( 1 );

		// remove session tokens label
		if ( isset( $user_meta['session_tokens'] ) ) {
			unset( $user_meta['session_tokens'] );
		}

		// Get all the keys, we'll use them as Column labels
		$user_meta_fields = array_keys( $user_meta );

		// Merge with BuddyPress labels if any
		$all_meta_labels = array_merge( $user_meta_fields, $bp_field_names );

		// Reset column letter offset, A-G reserved for basic user data
		$column_letter = 'K';

		// Set document properties
		$spreadsheet->getProperties()->setCreator( '' )
					->setLastModifiedBy( '' )
					->setTitle( 'Users' )
					->setSubject( 'all users' )
					->setDescription( 'Export of all users' )
					->setKeywords( 'office 2007 users export' )
					->setCategory( 'user results file' );

		// auto size column width
		foreach ( range( 'A', $spreadsheet->getActiveSheet()->getHighestDataColumn() ) as $col ) {
			$spreadsheet->getActiveSheet()
						->getColumnDimension( $col )
						->setAutoSize( true );
		}

		// Rename sheet
		$spreadsheet->getActiveSheet()->setTitle( 'Users' );

		// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$spreadsheet->setActiveSheetIndex( 0 );

		// Redirect output to a client’s web browser (Xlsx)
		header( 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
		header( 'Content-Disposition: attachment;filename="Users.xlsx"' );
		header( 'Cache-Control: max-age=0' );
		// If you're serving to IE 9, then the following may be needed
		header( 'Cache-Control: max-age=1' );

		// Save Excel file
		$writer = IOFactory::createWriter( $spreadsheet, 'Xlsx' );
		$writer->save( 'php://output' );
		exit;
	}

	return;
}

/**
 * @param $id
 * @param $fields
 * @param $consent
 *
 * @return array
 */
function get_from_users_table( $id, $fields, $consent ) {
	$data = [];

	$user_info        = get_userdata( $id );
	$forbidden        = [
		'user_pass',
	];
	$requires_consent = [
		'display_name',
		'first_name',
		'last_name',
		'user_email',
	];
	$requires_implode = [
		'roles',
	];

	foreach ( $fields as $k => $info ) {
		// protect PII
		if ( in_array( $info, $requires_consent, true ) ) {
			$info = ( $consent === '1' ) ? $user_info->data->$info : '';
		} elseif ( in_array( $info, $requires_implode, true ) ) {
			$info = implode( ', ', $user_info->$info );
		} elseif ( in_array( $info, $forbidden, true ) ) {
			$info = '';
		} else {
			$info = $user_info->data->$info;
		}

		$data[ $k ] = $info;
	}

	return $data;
}
