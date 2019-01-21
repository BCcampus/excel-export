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
							->SetCellValue( $column_letter . '1', esc_html( $label ) );
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
		} elseif ( ! empty( $_POST['export_posts'] ) ) {
			// try processing custom input
			$data    = [];
			$headers = [];
			$headers = apply_filters( 'excel_export_custom_data_headers', $headers );
			$data    = apply_filters( 'excel_export_custom_data', $data );

			// match the post data with the array key which defines the data
			if ( ! array_key_exists( $_POST['export_posts'], $headers ) ) {
				return;
			}

			$headers     = $headers[ $_POST['export_posts'] ];
			$spreadsheet = new Spreadsheet();
			$cell_count  = 1;
			$alphabet    = range( 'A', 'Z' );
			$num_columns = count( $headers );

			// get keys for only what we need
			$alpha_keys = array_splice( $alphabet, 0, $num_columns );

			// create cell headers from the alpha keys and filtered values
			$cell_headers = array_combine( $alpha_keys, array_values( $headers ) );

			/**
			 * Set cell headers
			 */
			if ( ! empty( $cell_headers ) ) {
				foreach ( $cell_headers as $k => $v ) {
					$spreadsheet->setActiveSheetIndex( 0 )
								->SetCellValue( $k . $cell_count, $v );
				}
			}

			foreach ( $data as $row ) {
				$cell_count++;

				$all_data_with_keys = array_combine( $alpha_keys, array_values( $row ) );

				// set csv data
				foreach ( $all_data_with_keys as $k => $v ) {
					$spreadsheet->setActiveSheetIndex( 0 )
								->SetCellValue( $k . $cell_count, $v );
				}
			}
			// Set document properties
			$spreadsheet->getProperties()->setCreator( '' )
						->setLastModifiedBy( '' )
						->setTitle( 'Custom' )
						->setSubject( 'custom' )
						->setDescription( 'Custom Export' )
						->setKeywords( 'office 2007 custom export' )
						->setCategory( 'custom results file' );

			// auto size column width
			foreach ( range( 'A', $spreadsheet->getActiveSheet()->getHighestDataColumn() ) as $col ) {
				$spreadsheet->getActiveSheet()
							->getColumnDimension( $col )
							->setAutoSize( true );
			}

			// Rename sheet
			$spreadsheet->getActiveSheet()->setTitle( 'Custom' );

			// Set active sheet index to the first sheet, so Excel opens this as the first sheet
			$spreadsheet->setActiveSheetIndex( 0 );

			// Redirect output to a client’s web browser (Xlsx)
			header( 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
			header( 'Content-Disposition: attachment;filename="Custom.xlsx"' );
			header( 'Cache-Control: max-age=0' );
			// If you're serving to IE 9, then the following may be needed
			header( 'Cache-Control: max-age=1' );

			// Save Excel file
			$writer = IOFactory::createWriter( $spreadsheet, 'Xlsx' );
			$writer->save( 'php://output' );
			exit;

		} else {
			$notice = 'Excel Export does not detect any data it can work with, please try again.';
		}
		?>
		<script type="text/javascript"><?php echo 'alert("' . $notice . '");'; ?></script>
		<?php
	}

}
