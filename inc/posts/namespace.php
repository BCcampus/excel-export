<?php

namespace BCcampus\Excel\Posts;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

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
