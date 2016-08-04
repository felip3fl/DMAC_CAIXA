<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" <?php language_attributes(); ?>>

<head profile="http://gmpg.org/xfn/11">
	<meta http-equiv="Content-Type" content="<?php bloginfo('html_type'); ?>; charset=<?php bloginfo('charset'); ?>" />
	
	<title><?php bloginfo('name'); ?> <?php if ( is_single() ) { ?> &raquo; Blog Archive <?php } ?> <?php wp_title(); ?></title>
	
	<meta name="generator" content="WordPress <?php bloginfo('version'); ?>" /> <!-- leave this for stats -->
	
	<link rel="stylesheet" href="<?php bloginfo('stylesheet_url'); ?>" type="text/css" media="screen" />
	<link rel="alternate" type="application/rss+xml" title="<?php bloginfo('name'); ?> RSS Feed" href="<?php bloginfo('rss2_url'); ?>" />
	<link rel="pingback" href="<?php bloginfo('pingback_url'); ?>" />
	
	<link rel="stylesheet" type="text/css" href="<?php bloginfo('template_directory'); ?>/style.css" />
	
	<!--[if IE]>
		<link rel="stylesheet" type="text/css" href="<?php bloginfo('template_directory'); ?>/style_ie.css" />
	<![endif]-->
	
	<!--[if lt IE 7]>
		<link rel="stylesheet" type="text/css" href="<?php bloginfo('template_directory'); ?>/style_ie6.css" />
	<![endif]-->
	
<?php wp_head(); ?>
</head>

<body>

<!-- Page -->
<div id="page">

	<!-- Header -->
	<div id="header">
	
		<!-- Menu -->
		<div id="header-menu">
			<ul>
				<?php wp_list_pages('title_li='); ?>
			</ul>
		</div>
		<!-- /Menu -->
		
		<!-- Title -->
		<div id="header-info">
			<h1><a href="<?php echo get_option('home'); ?>/"><?php bloginfo('name'); ?></a></h1>
			<div class="description"><?php bloginfo('description'); ?></div>
		</div>
		<!-- /Title -->
		
		<!-- Topbar -->
		<div id="topbar">
			<div id="topbar-left">
								
			</div>
			<div id="topbar-center">
				<h3></h3>
				<ul>
					<?php wp_get_archives('type=postbypost&limit=3'); ?>
				</ul>
			</div>
			<div id="topbar-right">
				<h3></h3>
				<ul>
					<?php
						$sql = "SELECT * FROM $wpdb->comments LEFT OUTER JOIN $wpdb->posts ON ($wpdb->comments.comment_post_ID = $wpdb->posts.ID) WHERE comment_approved = '1' AND comment_type = '' ORDER BY comment_date DESC LIMIT 0 , 3";
						$comments = $wpdb->get_results($sql);
						foreach ($comments as $comment) {
							$data = $comment->comment_author . " @ " . $comment->post_title;
							echo "<li><a href=\"" . get_permalink($comment->comment_post_ID) . "\">" . substr($data,0,40) . " ...</a></li>";   
						}
					?>
				</ul>
			</div>
		</div>
		<!-- /Topbar -->
	
	</div>
	<!-- /Header -->
	
	<!-- Main -->
	<div id="main"><div id="main-top"><div id="main-bottom">
	