<?php

namespace BCcampus\Excel\Users;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Gets and exports the user data
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

		$basic_info = [
			'ID' => 'User ID',
			'user_login' => 'Username',
			'display_name' => 'Display Name',
			'first_name' => 'First Name',
			'last_name' => 'Last Name',
			'user_email' => 'Email',
			'user_url' => 'URL',
			'user_registered' => 'Registration Date',
			'roles' => 'Roles',
			'user_level' => 'User Level',
			'user_status' => 'User Status',
		];

		apply_filters( 'excel_export_headers', $basic_info );

		$cell_headers = [
			'A' => 'User ID',
			'B' => 'Username',
			'C' => 'Display Name',
			'D' => 'First Name',
			'E' => 'Last Name',
			'F' => 'Email',
			'G' => 'URL',
			'H' => 'Registration Date',
			'I' => 'Roles',
			'J' => 'User Level',
			'K' => 'User Status',
		];

		// set csv headers
		foreach ( $cell_headers as $k => $v ) {
			$spreadsheet->setActiveSheetIndex( 0 )
						->SetCellValue( $k . $cell_count, $v );
		}

		// Get User Data and Meta for each user
		foreach ( $wp_users as $user ) {
			$cell_count ++;

			// Get basic user data
			$user_info = get_userdata( $user->ID );

			$basic = [
				'A' => $user_info->ID,
				'B' => $user_info->user_login,
				'C' => ( $consent === '1' ) ? $user_info->display_name : '',
				'D' => ( $consent === '1' ) ? $user_info->first_name : '',
				'E' => ( $consent === '1' ) ? $user_info->last_name : '',
				'F' => ( $consent === '1' ) ? $user_info->user_email : '',
				'G' => $user_info->user_url,
				'H' => $user_info->user_registered,
				'I' => implode( ', ', $user_info->roles ),
				'J' => $user_info->user_level,
				'K' => $user_info->user_status,
			];

			if ( function_exists( 'bp_is_active' ) ) {
				// Get the BP data for this user
				$bp_get_data = \BP_XProfile_ProfileData::get_data_for_user( $user_info->ID, $bp_field_ids );

				// Get the value of BP fields for this user
				foreach ( $bp_get_data as $bp_field_value ) {
					$bp_field_data [] = $bp_field_value->value;
				}
			}

			// set csv basic user data
			foreach ( $basic as $k => $v ) {
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

			// todo: find a way to add/map the data correctly regardless of what columns a user has
			/**
			 * // Merge with the BuddyPress data if any
			 * $all_meta = array_merge( $user_meta, $bp_field_data );
			 *
			 * // Add each user meta to appropriate excel column
			 * foreach ( $all_meta as $meta ) {
			 * $column_letter ++;
			 * $meta_value = is_serialized( $meta ); // check if it's serialized
			 * if (! $meta_value ) { // if unserialize() returns false, just get the meta value
			 * $meta_value = $meta; // get the meta value
			 * } else { // otherwise let's unserialized  the meta values
			 * $meta_value = maybe_unserialize($meta);
			 * $unserialized = [];
			 * foreach ( $meta_value as $key => $value ) {
			 * $unserialized[] = $key . ':' . $value;  // separate with a colon for readability
			 * }
			 * $meta_value = join( ', ', $unserialized ); // add comma separator for readability of multiple values
			 * }
			 * $spreadsheet->setActiveSheetIndex( 0 )
			 * ->SetCellValue( $column_letter . $cell_count, $meta_value ); // add meta value to the right column and cell
			 * }
			 */
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

		// Set up column labels for user meta
		// todo: find a way to add/map the data correctly regardless of what columns a user has
		/***
		 * foreach ( $all_meta_labels as $field ) {
		 * $column_letter ++;
		 * $spreadsheet->setActiveSheetIndex( 0 )
		 * ->SetCellValue( $column_letter . '1', $field );
		 * }
		 */

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
}
