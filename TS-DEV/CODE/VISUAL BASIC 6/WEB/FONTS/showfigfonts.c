/* showfigfonts for Windows 9x
   by Lionel Fourquaux [lionel.fourquaux@wanadoo.fr]
*/

#include <stdio.h>
#include <stdlib.h>
#include <io.h>
#include <process.h>
#include <string.h>
#include <stdarg.h>


void __declspec(noreturn) error(const int ret, const char * const msg, ...)
{
	va_list args;
	va_start(args, msg);
	vfprintf(stderr, msg, args);
	va_end(args);
	exit(ret);
}

int get_command_output(const char * const command, char * const buffer, unsigned int length)
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

void showfont(const char * const dirname, const unsigned int dirnamelen,
		const char * const fontname, const unsigned int fontnamelen,
		const char * const msg, const unsigned int msglen)
{
	char command[_MAX_PATH + _MAX_FNAME + 19];
	FILE *fp;
	memcpy(command, "figlet -d \"", 11);
	memcpy(command + 11, dirname, dirnamelen);
	memcpy(command + dirnamelen + 11, "\" -f \"", 6);
	memcpy(command + dirnamelen + 17, fontname, fontnamelen);
	memcpy(command + dirnamelen + fontnamelen + 17, "\"", 2);
	fflush(stdout);
	fp = _popen(command, "wt");
	if (fp == NULL)
		error(1, "Cannot execute %s\n", command);
	fwrite(msg, sizeof(char), msglen, fp);
	_pclose(fp);
}

void showallfonts(char * const dirname, unsigned int dirnamelen, const char * const msg)
{
	long fdh;
	struct _finddata_t fdinfo;
	unsigned int len, msglen;
	if (msg != NULL)
		msglen = strlen(msg);
	strcpy(dirname + dirnamelen, "\\*.flf");
	fdh = _findfirst(dirname, &fdinfo);
	if (fdh != -1)
	{
		do
		{
			len = strlen(fdinfo.name) - 4;
			fwrite(fdinfo.name, sizeof(char), len, stdout);
			fputs(":\n", stdout);
			if (msg == NULL)
				showfont(dirname, dirnamelen, fdinfo.name, len, fdinfo.name, len);
			else
				showfont(dirname, dirnamelen, fdinfo.name, len, msg, msglen);
			fputs("\n\n", stdout);
		}
		while (_findnext(fdh, &fdinfo) == 0);
		_findclose(fdh);
	};
}

int main(const int argc, const char * argv[])
{
	char dirname[_MAX_PATH];
	const char *msg;
	int dirnamelen;
	if (argc > 1 && strcmp(argv[1], "-d") == 0)
	{
		if (argc == 2 || argc > 4)
			goto usage;
		dirnamelen = strlen(argv[2]);
		memcpy(dirname, argv[2], dirnamelen);
		msg = (argc == 3) ? NULL : argv[3];
	}
	else
	{
		if (argc > 2)
			goto usage;
		dirnamelen = get_command_output("figlet -I2", dirname, _MAX_PATH);
		msg = (argc == 1) ? NULL : argv[1];
	};
	showallfonts(dirname, dirnamelen, msg);
	return 0;
usage:
	_splitpath(argv[0], NULL, NULL, dirname, NULL);
	error(2, "Usage: %s [ -d directory ] [ word ]", dirname);
}

