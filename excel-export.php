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

/**
 * Load dependencies
 */
include( __DIR__ . '/inc/users/namespace.php' );
include( __DIR__ . '/inc/posts/namespace.php' );


$composer = __DIR__ . '/vendor/autoload.php';
if ( file_exists( $composer ) ) {
	require $composer;
}

/**
 * Check permission levels, only proceed if we can manage_options
 */
add_action(
	'init', function () {
		if ( current_user_can( 'manage_options' ) ) {
			add_action( 'admin_menu', 'excel_export_admin_page' );
			add_action( 'admin_init', 'BCcampus\Excel\Users\excel_export_users' );
			add_action( 'admin_init', 'BCcampus\Excel\Posts\excel_export_posts' );
		} else {
			return;
		}
	}
);

/**
 * Add settings menu to the dashboard, and callback function for export page
 */
function excel_export_admin_page() {
	add_submenu_page( 'tools.php', 'Excel Export', 'Excel Export', 'manage_options', 'excel-export', 'excel_export_page' );
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

