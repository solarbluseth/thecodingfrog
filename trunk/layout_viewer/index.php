<!doctype html>
<?php
	
	$__maquettes = array();
	$dirname = './';
	$count = 0;
	foreach(glob($dirname.'*.jpg') as $file)
	{
		if($file != '.' && $file != '..' && !is_dir($dirname.$file))
		{
			$path_parts = pathinfo($dirname.$file);
			$__maquettes[$count]['file'] = $file;
			$__maquettes[$count]['filename'] = $path_parts['filename'];
			$__maquettes[$count]['filedate'] = filemtime($file);
			if ($_COOKIE[str_replace(".", "_", $path_parts['filename'])] != '')
			{
				if ($_COOKIE[str_replace(".", "_", $path_parts['filename'])] == filemtime($file))
				{
					$__maquettes[$count]['new'] = 0;
				}
				else
				{
					$__maquettes[$count]['new'] = 1;
				}
			}
			else
			{
				$__maquettes[$count]['new'] = 2;
			}
			SetCookie($path_parts['filename'], filemtime($file), time()+3600*24*365);
			$count++;
		}
	}

?>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <title>Prévisualisation des maquettes</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
	<link href="http://netdna.bootstrapcdn.com/font-awesome/4.1.0/css/font-awesome.min.css" rel="stylesheet">
  <script src="jquery.min.js"></script>
  <script type="text/javascript">
    $(document).ready(function(){
      $('body').on('click','.menu',function(e){
        e.preventDefault();
        $('aside').animate({
          'right':0
        });
        $('.bckg').fadeIn();
      });
      $('body').on('click','.bckg',function(){
        hideMenu();
      });
      $('body').on('click','aside ul li a',function(e){
        e.preventDefault();
        $("aside ul li").removeClass('active');
        $(this).parent("li").addClass('active');
        var path = $(this).attr('href');
        html = "<img src='"+path+"' title=''/>"
        $('.ctn').html(html);
        hideMenu();
      });
      $('body').on('click','.next',function(e){
        e.preventDefault();
        $("li.active").next("li").addClass('next');
        var path = $("li.next").children("a").attr('href');
        $("li.active").removeClass('active');
        if (path) {
          $("li.next").addClass('active').removeClass('next');
        }
        else {
          $("li").first().addClass('active');
          path = $('li').first().children('a').attr('href');
        }
        html = "<img src='"+path+"' title=''/>"
        $('.ctn').html(html);
      });
      $('body').on('click','.prev',function(e){
        e.preventDefault();
        $("li.active").prev("li").addClass('next');
        var path = $("li.next").children("a").attr('href');
        $("li.active").removeClass('active');
        if (path) {
          $("li.next").addClass('active').removeClass('next');
        }
        else {
          $("li.next").removeClass('next');
          $("li").last().addClass('active');
          path = $('li').last().children('a').attr('href');
        }
        html = "<img src='"+path+"' title=''/>"
        $('.ctn').html(html);
      });
    });
    function hideMenu() {
      $(".bckg").fadeOut();
      $('aside').animate({
        'right':'-85%'
      });
    }
  </script>
  <style>
    body,html{margin:0;font-size:14px;font-family:sans-serif;background-color:#000;}
    img{display:block;width:100%}
    .left{float:left}
    .right{float:right}
    .control{width:100%;position:fixed;top:0;left:0}
    .control a{width:33.33%;margin:0 auto;padding:50px 0;background-color:rgba(0,0,0,.0);color:rgba(255,255,255,.0);text-align:center;}
	.control a:hover{background-color:rgba(0,0,0,.5);color:rgba(255,255,255,1);}
    a.menu{width:33%;float:left;z-index: 2;text-align:center;}
    .bckg{display:none;width:100%;height:100%;position:fixed;z-index:1;top:0;left:0;background-color:rgba(0,0,0,.6)}
    aside{width:80%;position:fixed;z-index:3;top:0;right:-85%;text-align:left;height:100%;background-color:#444;box-shadow:-2px 0 10px rgba(0,0,0,.5);overflow-y:auto}
    ul{padding:0;list-style:none}
    //ul li{border-top:1px solid #383838}
    ul li a{padding:10px 10px;font-size:1.2em;color:#fff;text-decoration:none;//line-height:3;display:block;//border-top:1px solid #444;//border-top:1px solid #4E4E4E;}
    ul li.active a{color:#c0c0c0}
    ul li:first-child,ul li:first-child a{border:none}
    @media (min-width:1200px){
      body{max-width:320px;margin:0 auto}
    }
  </style>
</head>
<body>
   
  <div class="control">
    <a href="" class="right next"><i class="fa fa-angle-right fa-2x"></i></a>
    <a href="" class="left prev"><i class="fa fa-angle-left fa-2x"></i></a>
    <a href="#" class="menu"><i class="fa fa-bars fa-2x"></i></a>
  </div>

  <aside>
    <ul>
      <?php
        $dirname = './';
        $c = 0;
        $firstImg = "";
		
		for ($i = 0; $i < count($__maquettes); $i++)
		{
			echo "<li";
			if ($c == 0)
			{
				$firstImg = $dirname.$__maquettes[$i]["file"];
				echo " class='active'";
			}
			echo "><a href='".$dirname.$__maquettes[$i]["file"]."'>".$__maquettes[$i]['filename'];
			if ($__maquettes[$i]['new'] == 1) echo " <i style='font-size: 0.8em; color: #cdcdcd'>New</i>";
			echo "</a></li>";
			$c++;
		}
      ?>
    </ul>
  </aside>
  <div class="ctn">
    <?php
      echo "<img src='".$firstImg."' alt='".$firstImg."' />";
    ?>
  </div>
  <div class="bckg"></div>
</body>
</html>