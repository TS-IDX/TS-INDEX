/* figlist for Windows 9x
   by Lionel Fourquaux [lionel.fourquaux@wanadoo.fr]
*/

#include <stdio.h>
#include <stdlib.h>
#include <io.h>
#include <process.h>
#include <string.h>
#include <stdarg.h>
#include <errno.h>


void __declspec(noreturn) error(const int ret, const char * const msg, ...)
{
	va_list args;
	va_start(args, msg);
	vfprintf(stderr, msg, args);
	va_end(args);
	exit(ret);
}

int get_command_output(const char * const command, char * const buffer, const unsigned int length)
{
	FILE *fp;
	int count;
	fp = _popen(command, "rt");
	if (fp == NULL)
		error(1, "Cannot execute %s\n", command);
	count = fread(buffer, sizeof(char), length, fp);
	_pclose(fp);
	return count - 1;
}

void list_file_names(char * const dirname, const unsigned int dirnamelen,
		const char * const extension, const unsigned int extlen,
		const char * const foundmsg, const char * const notfoundmsg)
{
	long fdh;
	struct _finddata_t fdinfo;
	memcpy(dirname + dirnamelen + 2, extension, extlen + 1);
	fdh = _findfirst(dirname, &fdinfo);
	if (fdh == -1)
	{
		fputs(notfoundmsg, stdout);
	}
	else
	{
		fputs(foundmsg, stdout);
		do
		{
			fwrite(fdinfo.name, sizeof(char), strlen(fdinfo.name) - extlen, stdout);
			putchar('\n');
		}
		while (_findnext(fdh, &fdinfo) == 0);
		_findclose(fdh);
	};
}

int main(const int argc, const char * argv[])
{
	char deffontdir[_MAX_PATH], deffont[_MAX_FNAME];
	int deffontdirlen, deffontlen;
	if (argc == 1)
	{
		deffontdirlen = get_command_output("figlet -I2", deffontdir, _MAX_PATH);
	}
	else if (argc == 3 && strcmp(argv[1], "-d") == 0)
	{
		deffontdirlen = strlen(argv[2]);
		memcpy(deffontdir, argv[2], deffontdirlen);
	}
	else
	{
		_splitpath(argv[0], NULL, NULL, deffont, NULL);
		error(2, "Usage: %s [ -d directory ]", deffont);
	};
	deffontlen = get_command_output("figlet -I3", deffont, _MAX_FNAME);
	fputs("Default font: ", stdout);
	fwrite(deffont, sizeof(char), deffontlen, stdout);
	fputs("\nFont directory: ", stdout);
	fwrite(deffontdir, sizeof(char), deffontdirlen, stdout);
	putchar('\n');
	deffontdir[deffontdirlen] = '\0';
	if (_access(deffontdir, 4) < 0)
		error(1, "Cannot read %s: %s\n", deffontdir, strerror(errno));
	memcpy(deffontdir + deffontdirlen, "\\*", 2);
	list_file_names(deffontdir, deffontdirlen, ".flf", 4,
		"Figlet fonts in this directory:\n",
		"No figlet fonts in this directory\n");
	list_file_names(deffontdir, deffontdirlen, ".flc", 4,
		"Figlet control files in this directory:\n",
		"No figlet control files in this directory\n");
	return 0;
}

