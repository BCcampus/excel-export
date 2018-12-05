<?php
/**
 * Plugin Name:     Excel Export
 * Plugin URI:      https://github.com/BCcampus/excel-export
 * Description:     Export your posts, pages, custom post types, and user data to Excel file format (.XLSX)
 * Author:          Alex Paredes
 * Text Domain:     excel-export
 * Domain Path:     /languages
 * Version:         1.0.0
 *
 * @package         Excel_Export
 */

namespace BCcampus\Excel;

/**
 * Load dependencies
 */
$composer = __DIR__ . '/vendor/autoload.php';
if ( file_exists( $composer ) ) {
	require $composer;
}

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Check permission levels, only proceed if we can manage_options
 */

add_action( 'init', __NAMESPACE__ . '\excel_export_permissions' );

function excel_export_permissions() {
	if ( current_user_can( 'manage_options' ) ) {
		add_action( 'admin_menu', __NAMESPACE__ . '\excel_export_admin_page' );
		add_action( 'admin_init', __NAMESPACE__ . '\excel_export_users' );
		add_action( 'admin_init', __NAMESPACE__ . '\excel_export_posts' );
	} else {
		return;
	}
}

/**
 * Add settings menu to the dashboard, and callback function for export page
 */

function excel_export_admin_page() {
	add_submenu_page( 'tools.php', 'Excel Export', 'Excel Export', 'manage_options', 'excel-export', __NAMESPACE__ . '\excel_export_page' );
}

/**
 * Settings page content
 */

function excel_export_page() {
	// user count
	$user_count = count_users();
	// output buffering to capture output of echo into local var
	ob_start();
	foreach ( $user_count['avail_roles'] as $role => $count ) {
		echo ' ', $count, ' are ', $role, ',';
	}
	$role_count = rtrim( ob_get_contents(), ',' );
	ob_end_clean();

	// get the post types on this website
	$args = [
		'public' => true,
	];

	$output     = 'names';
	$operator   = 'and';
	$post_types = get_post_types( $args, $output, $operator );

	// page content
	$html  = '<form action="#post-export" method="post">';
	$html .= '<p><h1>Excel Export<span class="dashicons dashicons-download"></span></h1></p>';
	$html .= '<hr><p><h2>Export Post Types</h2><p>The following post types were found on your website: </p>';
	$html .= '<select id="excel_export_users" name="export_posts" />';
	// let's populate the select list from the post types available on this website
	foreach ( $post_types as $post_type ) {
		$html .= '<option value="' . esc_attr( $post_type ) . '">' . esc_attr( $post_type ) . '</option>';
	}
	// post export button
	$html .= '</select><input class="button button-primary export_button" style="margin-top:3px;" type="submit" id="excel_export_posts_submit" name="export_posts_submit" value="Export" /></p>';
	// post export nonce
	$html .= wp_nonce_field( 'export_button_posts', 'submit_export_posts' );
	$html .= '</form>';
	echo $html;

	// user export button
	$html = '<form action="#user-export" method="post">';
	// user export nonce
	$html .= '<hr><p><h2>Export Users</h2></p>There are <u>' . esc_attr( $user_count['total_users'] ) . '</u> users in total:' . esc_attr( $role_count ) . '. </p><input class="button button-primary export_button" style="margin-top:3px;" type="submit" id="excel_export_users" name="users_export" value="Export Users" /></p>';
	$html .= '<input type="checkbox" name="consent" value="1"> Include personally identifiable information in export<hr>';
	$html .= wp_nonce_field( 'export_button_users', 'submit_export_users' );
	$html .= '</form>';
	echo $html;
}

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

		apply_filters( 'excel_export_headers', $cell_headers );

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

/**
 * Gets and exports the post data
 */

function excel_export_posts() {
	// check if Post data is being requested and nonce is valid
	if ( ! empty( $_POST ) && isset( $_POST['export_posts'] ) && check_admin_referer( 'export_button_posts', 'submit_export_posts' ) ) {

		// Create a new PHPExcel object
		$spreadsheet = new Spreadsheet();

		$post_type_requested = $_POST['export_posts'];

		if ( post_type_exists( $post_type_requested ) ) {

			// post args
			$args = [
				'post_type'      => $post_type_requested,
				'posts_per_page' => - 1,
				'offset'         => 0,
			];

			// post query
			$posts = get_posts( $args );

			// Initial count for rows
			$count = 1;

			// Get the data we want from each post
			foreach ( $posts as $single ) {
				// Set initial column letter
				$column_letter = 'A';

				$count ++;

				foreach ( $single as $meta ) {
					$post_labels = [];
					$post_values = [];
					foreach ( $single as $key => $value ) {
						$post_labels[] = $key;
						$post_values[] = $value;
					}
				}
				// Set up column values for post meta
				foreach ( $post_values as $val ) {
					$spreadsheet->setActiveSheetIndex( 0 )
								->SetCellValue( $column_letter . $count, $val );
					$column_letter ++;
				}
			}

			// Reset the column letter
			$column_letter = 'A';
			// Set up column labels for post meta
			foreach ( $post_labels as $label ) {
				$spreadsheet->setActiveSheetIndex( 0 )
							->SetCellValue( $column_letter . '1', esc_html__( $label ) );
				$column_letter ++;
			}

			// current blog time for the export name
			$blogtime = current_time( '--M-D-Y--H-I-s' );

			// Set document properties
			$spreadsheet->getProperties()->setTitle( esc_html( $post_type_requested ) );
			$spreadsheet->getProperties()->setSubject( esc_html( 'all ' . $post_type_requested ) );
			$spreadsheet->getProperties()->setDescription( esc_html( 'Export of all ' . $post_type_requested ) );

			// Rename sheet
			$spreadsheet->getActiveSheet()->setTitle( esc_html( $post_type_requested ) );

			// Rename file
			header( 'Content-Disposition: attachment;filename="' . $post_type_requested . $blogtime . '.xlsx"' );

			// Set column data auto width
			for ( $col = 'A'; $col !== 'E'; $col ++ ) {
				$spreadsheet->getActiveSheet()->getColumnDimension( $col )->setAutoSize( true );
			}

			header( 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
			header( 'Cache-Control: max-age=0' );

			// Save Excel file
			$obj_writer = IOFactory::createWriter( $spreadsheet, 'Xlsx' );
			$obj_writer->save( 'php://output' );

			exit();
		} else { // in the unlikely event an empty or invalid post type value is entered, let's display an ugly error
			$post_value = $_POST['export_posts'];
			if ( $post_value === '' ) {
				$notice = __( 'Export Error: Please select a post type to export it.', 'excel-export' );
			} else {
				$notice = 'Excel Export: ' . $post_value . ' does not exist, please try a different post type.';
			}
			?>
			<script type="text/javascript"><?php echo 'alert("' . $notice . '");'; ?></script>
			<?php
		}
	}
}
