CREATE TABLE IF NOT EXISTS `pm` (
`id` int(11) NOT NULL,
  `user_name` varchar(60) NOT NULL,
  `user_email` varchar(60) NOT NULL,
  `user_message` text NOT NULL
)AUTO_INCREMENT=1 ;

ALTER TABLE `pm`
 ADD PRIMARY KEY (`id`);

ALTER TABLE `pm`
MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;