CREATE TABLE  `since_forum_main` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `name` text NOT NULL,
  `lastposter` varchar(200) NOT NULL,
  `lastpostdate` int(11) NOT NULL,
  `topics` int(11) NOT NULL default '0',
  `replies` int(11) NOT NULL default '0',
  PRIMARY KEY  (`id`)
)

CREATE TABLE `since_forum_replies` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `topicid` int(11) unsigned,
  `message` text NOT NULL,
  `subject` text NOT NULL,
  `poster` text NOT NULL,
  `date` int(10) unsigned NOT NULL,
  PRIMARY KEY  (`id`)
)

CREATE TABLE  `since_forum_topics` (
  `id` int(11) unsigned NOT NULL auto_increment,
  `forumid` int(11) unsigned,
  `message` text NOT NULL,
  `subject` text NOT NULL,
  `poster` text NOT NULL,
  `date` int(11) NOT NULL,
  `lastposter` varchar(200) NOT NULL,
  `lastpostdate` int(11) NOT NULL,
  `replies` int(11) unsigned NOT NULL,
  PRIMARY KEY  (`id`)
)