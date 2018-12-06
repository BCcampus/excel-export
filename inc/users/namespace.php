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
		$spreadsheet       = new Spreadsheet();
		$args              = [
			'order'   => 'ASC',
			'orderby' => 'display_name',
			'fields'  => 'all',
		];
		$wp_users          = get_users( $args );
		$cell_count        = 1;
		$consent           = ( isset( $_POST['users_export'] ) ) ? $_POST['consent'] : 0;
		$alphabet          = range( 'A', 'Z' );
		$default_user_data = [
			'ID'              => 'User ID',
			'user_login'      => 'Username',
			'display_name'    => 'Display Name',
			'user_email'      => 'Email',
			'user_url'        => 'URL',
			'user_registered' => 'Registration Date',
			'roles'           => 'Roles',
			'user_level'      => 'User Level',
			'user_status'     => 'User Status',
		];

		// optional
		$default_user_meta       = [];
		$default_user_buddypress = [];

		// Get BuddyPress profile field ID's and names
		$bp_fields = get_bp_fields();

		// hooks
		$default_user_data       = apply_filters( 'excel_export_user_data', $default_user_data );
		$default_user_meta       = apply_filters( 'excel_export_user_metadata', $default_user_meta );
		$default_user_buddypress = apply_filters( 'excel_export_user_buddypress', $default_user_buddypress );

		// combine user and user_meta arrays
		$combined = array_merge( $default_user_data, $default_user_meta );
		$combined = array_merge( $combined, $default_user_buddypress );

		$num_columns = count( $combined );

		// get keys for only what we need
		$alpha_keys = array_splice( $alphabet, 0, $num_columns );

		// create cell headers from the alpha keys and filtered values
		$cell_headers = array_combine( $alpha_keys, array_values( $combined ) );

		/**
		 * Set cell headers
		 */
		foreach ( $cell_headers as $k => $v ) {
			$spreadsheet->setActiveSheetIndex( 0 )
						->SetCellValue( $k . $cell_count, $v );
		}

		/**
		 * Get User data for each user
		 */
		foreach ( $wp_users as $user ) {
			$cell_count ++;

			$user_content = get_from_users_table( $user->ID, array_keys( $default_user_data ), $consent );
			$user_meta    = get_from_usermeta_table( $user->ID, $default_user_meta );
			$user_buddy   = get_from_user_buddypress( $user->ID, $default_user_buddypress, $bp_fields );

			// combine
			$all_data = array_merge( $user_content, $user_meta );
			$all_data = array_merge( $all_data, $user_buddy );

			// give array combination alphabetic key values
			$all_data_with_keys = array_combine( $alpha_keys, array_values( $all_data ) );

			// set csv basic user data
			foreach ( $all_data_with_keys as $k => $v ) {
				$spreadsheet->setActiveSheetIndex( 0 )
							->SetCellValue( $k . $cell_count, $v );
			}
		}

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
 * @return array
 */
function get_bp_fields() {
	if ( ! function_exists( 'bp_is_active' ) ) {
		return [];
	}

	$bp_fields = [];

	$profile_groups = \BP_XProfile_Group::get(
		[
			'fetch_fields' => true,
		]
	);

	if ( ! empty( $profile_groups ) ) {
		foreach ( $profile_groups as $profile_group ) {
			if ( ! empty( $profile_group->fields ) ) {
				foreach ( $profile_group->fields as $field ) {
					$bp_fields[ $field->id ] = $field->name;
				}
			}
		}
	}

	return $bp_fields;
}

/**
 * @param $id
 * @param $fields
 * @param $bp_ids
 *
 * @return array
 */
function get_from_user_buddypress( $id, $fields, $bp_ids ) {
	if ( empty( $fields ) ) {
		return [];
	}
	$field_ids = [];
	$data      = [];
	$bp        = [];
	// ensures result length is the same coming in
	// as going out
	foreach ( $fields as $k => $v ) {
		$data[ $k ] = $v;
	}

	foreach ( $bp_ids as $k => $v ) {
		if ( array_key_exists( $v, $fields ) ) {
			$field_ids[] = $k;
		}
	}

	if ( function_exists( 'bp_is_active' ) && ! empty( $field_ids ) ) {
		// Get the BP data for this user
		$bp_get_data = \BP_XProfile_ProfileData::get_data_for_user( $id, $field_ids );

		// Get the value of BP fields for this user
		foreach ( $bp_get_data as $bp_field_value ) {
			$bp[] = $bp_field_value->value;
		}
	}

	$data = ( false !== array_combine( $data, $bp ) ) ? array_combine( $data, $bp ) : [];

	return array_values( $data );
}

/**
 * @param $id
 * @param $fields
 * @param $consent
 *
 * @return array
 */
function get_from_users_table( $id, $fields, $consent ) {
	if ( empty( $fields ) ) {
		return [];
	}

	$data = [];

	$user_info        = get_userdata( $id );
	$forbidden        = [
		'user_pass',
	];
	$requires_consent = [
		'display_name',
		'user_email',
	];
	$requires_implode = [
		'roles',
	];

	foreach ( $fields as $info ) {
		// protect PII
		if ( in_array( $info, $requires_consent, true ) ) {
			$info = ( $consent === '1' ) ? $user_info->data->$info : '';
			// deal with arrays
		} elseif ( in_array( $info, $requires_implode, true ) ) {
			$info = implode( ', ', $user_info->$info );
			// forbid certain data types
		} elseif ( in_array( $info, $forbidden, true ) ) {
			$info = '';
		} else {
			$info = $user_info->data->$info;
		}

		$data[] = $info;
	}

	return $data;
}


/**
 * @param $id
 * @param $fields
 *
 * @return array
 */
function get_from_usermeta_table( $id, $fields ) {
	if ( empty( $fields ) ) {
		return [];
	}
	$data = [];

	// ensures result length is the same coming in
	// as going out
	foreach ( $fields as $k => $v ) {
		$data[ $k ] = $v;
	}

	$forbidden = [
		'session_tokens',
	];

	// Get all the user meta into an array, run array_map to take only the first index of each result
	$user_meta = array_map(
		function ( $a ) {
			return $a[0];
		}, get_user_meta( $id )
	);

	foreach ( $fields as $k => $v ) {
		if ( isset( $user_meta[ $k ] ) ) {
			$data[ $k ] = maybe_unserialize( $user_meta[ $k ] );
			if ( is_array( $data[ $k ] ) ) {
				$csv        = implode( ', ', $data[ $k ] );
				$data[ $k ] = $csv;
			}
		}
	}

	// remove information with potential security implications
	foreach ( $forbidden as $remove ) {
		if ( array_key_exists( $remove, $data ) ) {
			unset( $data[ $remove ] );
		}
	}

	return array_values( $data );
}
