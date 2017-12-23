<?php
/**
 * Plugin Name:     Excel Export
 * Plugin URI:      https://github.com/BCcampus/excel-export
 * Description:     Export event and user data
 * Author:          Alex Paredes
 * Text Domain:     excel-export
 * Domain Path:     /languages
 * Version:         0.1.0
 *
 * @package         Excel_Export
 */

namespace BCcampus\Excel;

/**
 * Load dependencies
 */
require_once __DIR__ . '/vendor/autoload.php';

/**
 * Adds the 'Export to Excel' button on users and events admin screens
 */
function export_button() {
	// only show the export button if PHPExcel exists
	if ( class_exists( 'PHPExcel' ) ) {
		// add export button only on the event and users screen
		$screen      = get_current_screen();
		$allowed     = array( 'edit-event', 'users' );
		$unique_name = '';
		if ( ! in_array( $screen->id, $allowed ) ) {
			return;
		}
		if ( $screen->id == 'users' ) {
			$unique_name = 'users_export';
		} elseif ( $screen->id == 'edit-event' ) {
			$unique_name = 'events_export';
		}
		?>
        <script type="text/javascript">
            jQuery(document).ready(function ($) {
                $('.tablenav.top .clear, .tablenav.bottom .clear').before('<form action="#" method="POST"><input type="hidden" id="wp_excel_export" name="<?php echo $unique_name; ?>" value="1" /><input class="button button-primary export_button" style="margin-top:3px;" type="submit" value="<?php esc_attr_e( 'Export to Excel' );?>" /></form>');
            });
        </script>
		<?php
	}
}

add_action( 'admin_footer', __NAMESPACE__ . '\export_button' );

/**
 * Gets and exports the user and event data
 */
function export() {

	if ( ! empty( $_POST['users_export'] ) || ! empty( $_POST['events_export'] ) ) {

		if ( current_user_can( 'manage_options' ) ) {

			// Create a new PHPExcel object
			$objPHPExcel = new \PHPExcel();

			// Check if User data is being requested
			if ( isset( $_POST['users_export'] ) ) {

				// Args for the user query
				$args = array(
					'order'   => 'ASC',
					'orderby' => 'display_name',
					'fields'  => 'all',
				);

				// User Query
				$wp_users   = get_users( $args );
				$cell_count = 1;

				// BuddyPress user data
				$bp_field_names = array();
				$bp_field_ids   = array();

				// Get BuddyPress profile data if available
				if ( function_exists( 'bp_is_active' ) ) {

					$profile_groups = \BP_XProfile_Group::get( array( 'fetch_fields' => true ) );

					if ( ! empty( $profile_groups ) ) {
						foreach ( $profile_groups as $profile_group ) {
							if ( ! empty( $profile_group->fields ) ) {
								foreach ( $profile_group->fields as $field ) {
									$bp_field_names[] = $field->name;
									$bp_field_ids[]   = $field->id; // get field value by using BP_XProfile_ProfileData::get_value_byid()
								}
							}
						}
					}
				}

				// Get User Data and Meta for each user
				foreach ( $wp_users as $user ) {
					$cell_count ++;

					// Get basic user data
					$user_info    = get_userdata( $user->ID );
					$id           = $user_info->ID;
					$username     = $user_info->user_login;
					$email        = $user_info->user_email;
					$url          = $user_info->user_url;
					$registered   = $user_info->user_registered;
					$login        = $user_info->user_login;
					$display_name = $user_info->display_name;

					// Get all the meta, run array_map to take only the first index of each result
					$user_meta = array_map( function ( $a ) {
						return $a[0];
					}, get_user_meta( $user->ID ) );


					// Add basic user data to appropriate column
					$objPHPExcel->setActiveSheetIndex( 0 );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'A' . $cell_count . '', $id );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'B' . $cell_count . '', $username );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'C' . $cell_count . '', $email );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'D' . $cell_count . '', $url );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'E' . $cell_count . '', $registered );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'F' . $cell_count . '', $login );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'G' . $cell_count . '', $display_name );

					// Offset column letter, A-G reserved for basic user data
					$column_letter = 'G';

					// Add all user meta to appropriate column
					foreach ( $user_meta as $meta ) {
						$column_letter ++;
						$user_meta[ $meta ] += 1;
						$objPHPExcel->getActiveSheet()->SetCellValue( $column_letter . $cell_count . '', $meta );
					}

				}

				// user_id 1 as a placeholder to get column labels
				$user_meta = get_user_meta( 1 );

				// Get all the keys, we'll use them as Column labels
				$user_meta_fields = array_keys( $user_meta );

				// Reset column letter offset, A-G reserved for basic user data
				$column_letter = 'G';

				// Set up column labels for basic user data
				$objPHPExcel->getActiveSheet()->SetCellValue( 'A1', esc_html__( 'User ID' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'B1', esc_html__( 'Username' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'C1', esc_html__( 'Email' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'D1', esc_html__( 'URL' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'E1', esc_html__( 'Registration Date' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'F1', esc_html__( 'Login' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'G1', esc_html__( 'Display Name' ) );

				/*				// Set up column labels for Buddy Press
								foreach ( $bp_fields as $field ) {
									$column_letter ++;
									$bp_fields[ $field ] += 1;
									$objPHPExcel->getActiveSheet()->SetCellValue( $column_letter . '1', $field );
								}
				*/

				// Set up column labels for user meta
				foreach ( $user_meta_fields as $field ) {
					$column_letter ++;
					$user_meta_fields[ $field ] += 1;
					$objPHPExcel->getActiveSheet()->SetCellValue( $column_letter . '1', $field );
				}

				// Set document properties
				$objPHPExcel->getProperties()->setTitle( esc_html__( 'Users' ) );
				$objPHPExcel->getProperties()->setSubject( esc_html__( 'all users' ) );
				$objPHPExcel->getProperties()->setDescription( esc_html__( 'Export of all users' ) );

				// Rename sheet
				$objPHPExcel->getActiveSheet()->setTitle( esc_html__( 'Users' ) );

				// Rename file
				header( 'Content-Disposition: attachment;filename="users.xlsx"' );

			}

			// Event data
			if ( isset( $_POST['events_export'] ) ) {

				// Event args
				$args = array(
					'post_type'      => 'event',
					'posts_per_page' => - 1,
					'offset'         => 0,
				);

				// Event Query
				$posts      = get_posts( $args );
				$cell_count = 1;

				// Set up column labels
				$objPHPExcel->getActiveSheet()->SetCellValue( 'A1', esc_html__( 'Event Title' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'B1', esc_html__( 'Owner' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'C1', esc_html__( 'Status' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'D1', esc_html__( 'Published date' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'E1', esc_html__( 'Start date' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'F1', esc_html__( 'End date' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'G1', esc_html__( 'Start time' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'H1', esc_html__( 'End time' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'I1', esc_html__( 'Presenter' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'J1', esc_html__( 'Registration Contact E-mail' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'K1', esc_html__( 'Location ID' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'L1', esc_html__( 'Event ID' ) );

				// Get the data we want from each event
				foreach ( $posts as $post ) {
					$cell_count ++;

					$title       = $post->post_title;
					$author_id   = $post->post_author;
					$author      = get_the_author_meta( 'display_name', $author_id );
					$status      = $post->post_status;
					$date_pub    = $post->post_date;
					$start       = $post->_event_start_date;
					$end         = $post->_event_end_date;
					$start_time  = $post->_event_start_time;
					$end_time    = $post->_event_end_time;
					$presenter   = $post->{'Presenter(s)'};
					$reg_email   = $post->{'Registration Contact Email'};
					$location_id = $post->_location_id;
					$event_id    = $post->_event_id;

					// Add the event data to the appropriate column
					$objPHPExcel->getActiveSheet()->SetCellValue( 'A' . $cell_count . '', $title );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'B' . $cell_count . '', $author );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'C' . $cell_count . '', $status );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'D' . $cell_count . '', $date_pub );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'E' . $cell_count . '', $start );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'F' . $cell_count . '', $end );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'G' . $cell_count . '', $start_time );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'H' . $cell_count . '', $end_time );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'I' . $cell_count . '', $presenter );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'J' . $cell_count . '', $reg_email );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'K' . $cell_count . '', $location_id );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'L' . $cell_count . '', $event_id );

				}

				// Set document properties
				$objPHPExcel->getProperties()->setTitle( esc_html__( 'Events' ) );
				$objPHPExcel->getProperties()->setSubject( esc_html__( 'all events' ) );
				$objPHPExcel->getProperties()->setDescription( esc_html__( 'Export of all events' ) );

				// Rename sheet
				$objPHPExcel->getActiveSheet()->setTitle( esc_html__( 'Events' ) );

				// Rename file
				header( 'Content-Disposition: attachment;filename="events.xlsx"' );

			}

			// Set column data auto width
			for ( $col = 'A'; $col !== 'E'; $col ++ ) {
				$objPHPExcel->getActiveSheet()->getColumnDimension( $col )->setAutoSize( true );
			}
		}

		header( 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
		header( 'Cache-Control: max-age=0' );

		// Save Excel file
		$objWriter = \PHPExcel_IOFactory::createWriter( $objPHPExcel, 'Excel2007' );
		$objWriter->save( 'php://output' );

		exit();
	}
}

add_action( 'admin_init', __NAMESPACE__ . '\export' );
