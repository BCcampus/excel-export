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

			// User data
			if ( isset( $_POST['users_export'] ) ) {

				// User args
				$args = array(
					'order'   => 'ASC',
					'orderby' => 'display_name',
					'fields'  => 'all',
				);

				// User Query
				$wp_users   = get_users( $args );
				$cell_count = 1;

				// Set up column labels
				$objPHPExcel->getActiveSheet()->SetCellValue( 'A1', esc_html__( 'First Name' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'B1', esc_html__( 'Last Name' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'C1', esc_html__( 'Email' ) );
				$objPHPExcel->getActiveSheet()->SetCellValue( 'D1', esc_html__( 'User Role' ) );

				// Get the data we want from each user
				foreach ( $wp_users as $user ) {
					$cell_count ++;

					$user_meta  = get_user_meta( $user->ID );
					$role       = implode( ',', $user->roles );
					$email      = $user->user_email;
					$first_name = ( isset( $user_meta['first_name'][0] ) && $user_meta['first_name'][0] != '' ) ? $user_meta['first_name'][0] : '';
					$last_name  = ( isset( $user_meta['last_name'][0] ) && $user_meta['last_name'][0] != '' ) ? $user_meta['last_name'][0] : '';

					// Add the user data to the appropriate column
					$objPHPExcel->setActiveSheetIndex( 0 );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'A' . $cell_count . '', $first_name );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'B' . $cell_count . '', $last_name );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'C' . $cell_count . '', $email );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'D' . $cell_count . '', $role );

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

				// Get the data we want from each event
				foreach ( $posts as $post ) {
					$cell_count ++;

					$title     = $post->post_title;
					$date      = $post->post_date;
					$status    = $post->post_status;
					$author_id = $post->post_author;;
					$author   = get_the_author_meta( 'display_name', $author_id );
					$location = get_post_meta( $post->ID, 'location', true );

					// Add the event data to the appropriate column
					$objPHPExcel->getActiveSheet()->SetCellValue( 'A' . $cell_count . '', $title );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'B' . $cell_count . '', $author );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'C' . $cell_count . '', $status );
					$objPHPExcel->getActiveSheet()->SetCellValue( 'D' . $cell_count . '', $date );

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
