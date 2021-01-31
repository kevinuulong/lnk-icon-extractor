'use-strict';
const shell = require('node-powershell');
const path = require('path');

function registerShell() {
	return new shell({
		executionPolicy: 'Bypass',
		noProfile: true
	});
}

exports.extract = function (filePath, destinationPath, format = "png") {
	const filePaths = filePath => [].concat(filePath),
		ps = registerShell();
	ps.addCommand(`Add-Type -AssemblyName System.Drawing`)
	filePaths(filePath).forEach(element => {
		ps.addCommand(`$sh = New-Object -ComObject WScript.Shell`);
		ps.addCommand(`$exeLocation = $sh.CreateShortcut('${element}').TargetPath`);
		var finalDestinationPath = path.join(destinationPath, `${path.basename(element, path.extname(element))}.${format}`);
		ps.addCommand(`[System.Drawing.Icon]::ExtractAssociatedIcon($exeLocation).ToBitmap().Save('${finalDestinationPath}','${format}')`)
	});
	ps.invoke().then().catch(err => {
		console.log(err);
		ps.dispose();
	});
}
