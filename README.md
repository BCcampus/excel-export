# excel-export
[![Build Status](https://travis-ci.com/BCcampus/excel-export.svg?branch=dev)](https://travis-ci.com/BCcampus/excel-export)

Export your post, page, custom post type, and user data to Excel file format (.XLSX) 

## Description ##

Get all of your WordPress data into Excel (.XLSX file). This plugin allows you to export all of your post data, including custom post types. And all User data, including BuddyPress profile fields. 

## Features ##
- Export any Post Type. Finds all available post types on your website for export. 
- Export all User Data. Includes custom fields, custom meta, and BuddyPress profile fields. 

## Installation ##

1. Upload `excel-export.php` to the `/wp-content/plugins/` directory.
2. Activate the plugin through the 'Plugins' menu in WordPress.
3. Click on the "Excel Export" menu item located under the Tools menu.

## BuddyPress Extended profile data ##
By default, extended profile data is not included in your export. To export this data please use [add_filter](https://developer.wordpress.org/reference/functions/add_filter/) as show below:
```
add_filter( 'excel_export_user_buddypress', function ( $default_user_buddypress ) {
	$add_buddypress = [
		'YourField'                  => 'YourField',
		'AnotherField'               => 'AnotherField',
	];
	return array_merge( $default_user_buddypress, $add_buddypress );
} );
```


## Screenshots ##

![Export Button](/assets/img/settings.png)
![Export Button](/assets/img/menu_item.png)
