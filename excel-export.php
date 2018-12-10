<?php
/**
 * Plugin Name:     Excel Export
 * Plugin URI:      https://github.com/BCcampus/excel-export
 * Description:     Export your posts, pages, custom post types, and user data to Excel file format (.XLSX)
 * Author:          Alex Paredes
 * Text Domain:     excel-export
 * Domain Path:     /languages
 * Version:         0.5.0
 *
 * @package         Excel_Export
 */

/**
 * Load dependencies
 */

$composer = __DIR__ . '/vendor/autoload.php';
if ( file_exists( $composer ) ) {
	require $composer;
}

include( __DIR__ . '/inc/admin/namespace.php' );
include( __DIR__ . '/inc/users/namespace.php' );
include( __DIR__ . '/inc/posts/namespace.php' );

/**
 * Check permission levels, only proceed if we can manage_options
 */
add_action(
	'init', function () {
		if ( current_user_can( 'manage_options' ) ) {
			add_action( 'admin_menu', 'BCcampus\Excel\Admin\excel_export_admin_page' );
			add_action( 'admin_init', 'BCcampus\Excel\Users\excel_export_users' );
			add_action( 'admin_init', 'BCcampus\Excel\Posts\excel_export_posts' );
		} else {
			return;
		}
	}
);


