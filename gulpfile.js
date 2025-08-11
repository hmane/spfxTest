// gulpfile.js
const gulp = require('gulp');
const sass = require('gulp-sass')(require('sass'));
const sourcemaps = require('gulp-sourcemaps');
const postcss = require('gulp-postcss');
const autoprefixer = require('autoprefixer');
const path = require('path');
const fs = require('fs');

// Paths configuration
const paths = {
	styles: {
		src: 'src/**/*.scss',
		moduleSrc: 'src/**/*.module.scss',
		dest: 'lib',
		watch: 'src/**/*.scss',
	},
};

// PostCSS plugins
const postcssPlugins = [
	autoprefixer({
		overrideBrowserslist: ['>0.2%', 'not dead', 'not ie <= 11', 'not op_mini all'],
	}),
];

// Sass options
const sassOptions = {
	outputStyle: 'compressed',
	includePaths: ['node_modules', 'src'],
	quietDeps: true,
};

const sassModuleOptions = {
	outputStyle: 'expanded',
	includePaths: ['node_modules', 'src'],
	quietDeps: true,
};

// Generate TypeScript definitions for SCSS modules
function generateScssTypings(cssContent, originalPath) {
	try {
		const classNames = [];
		const regex = /\.([a-zA-Z_][\w-]*)/g;
		let match;

		while ((match = regex.exec(cssContent)) !== null) {
			const className = match[1];
			if (!classNames.includes(className)) {
				classNames.push(className);
			}
		}

		const typings = `// This file is auto-generated. Do not edit manually.
declare const styles: {
${classNames.map((name) => `  readonly ${JSON.stringify(name)}: string;`).join('\n')}
};
export default styles;
`;

		const dtsPath = originalPath.replace('.scss', '.scss.d.ts');
		fs.writeFileSync(dtsPath, typings);
		console.log(`Generated typings for: ${path.relative(process.cwd(), dtsPath)}`);
	} catch (error) {
		console.warn('Failed to generate typings:', error.message);
	}
}

// Custom error handler
function handleSassError(error) {
	if (
		error.message &&
		(error.message.includes('@import') ||
			error.message.includes('deprecated') ||
			error.message.includes('will be removed'))
	) {
		console.log(`SCSS Warning: ${error.message.split('\n')[0]}`);
		this.emit('end');
	} else {
		console.error('SCSS Error:', error.message);
		this.emit('end');
	}
}

// Task to build regular SCSS files (non-module)
function buildRegularStyles() {
	return gulp
		.src([paths.styles.src, `!${paths.styles.moduleSrc}`])
		.pipe(sourcemaps.init())
		.pipe(sass(sassOptions).on('error', handleSassError))
		.pipe(postcss(postcssPlugins))
		.pipe(sourcemaps.write('.'))
		.pipe(gulp.dest(paths.styles.dest));
}

// Task to build SCSS modules
function buildModuleStyles() {
	return gulp
		.src(paths.styles.moduleSrc)
		.pipe(sourcemaps.init())
		.pipe(sass(sassModuleOptions).on('error', handleSassError))
		.pipe(postcss(postcssPlugins))
		.pipe(sourcemaps.write('.'))
		.pipe(gulp.dest(paths.styles.dest))
		.on('data', function (file) {
			if (file.path.includes('.module.css')) {
				try {
					const originalScssPath = file.path
						.replace(path.resolve('lib'), path.resolve('src'))
						.replace('.css', '.scss');
					generateScssTypings(file.contents.toString(), originalScssPath);
				} catch (error) {
					console.warn('Failed to process module file:', error.message);
				}
			}
		});
}

// Task to copy SCSS files
function copyScssFiles() {
	return gulp.src(paths.styles.src).pipe(gulp.dest(paths.styles.dest));
}

// Task to generate global TypeScript definitions
function generateGlobalTypings() {
	const globalTypings = `// Global SCSS module declarations
declare module '*.scss' {
  const content: { [className: string]: string };
  export default content;
}

declare module '*.module.scss' {
  const classes: { [key: string]: string };
  export default classes;
}

declare module '*.sass' {
  const content: { [className: string]: string };
  export default content;
}

declare module '*.module.sass' {
  const classes: { [key: string]: string };
  export default classes;
}

declare module '*.css' {
  const content: { [className: string]: string };
  export default content;
}

declare module '*.module.css' {
  const classes: { [key: string]: string };
  export default classes;
}
`;

	try {
		if (!fs.existsSync('lib')) {
			fs.mkdirSync('lib', { recursive: true });
		}
		fs.writeFileSync('lib/scss.d.ts', globalTypings);

		if (!fs.existsSync('src/types')) {
			fs.mkdirSync('src/types', { recursive: true });
		}
		fs.writeFileSync('src/types/scss.d.ts', globalTypings);

		console.log('Generated global SCSS typings');
	} catch (error) {
		console.warn('Failed to generate global typings:', error.message);
	}

	return Promise.resolve();
}

// Clean task
function cleanStyles() {
	try {
		const { execSync } = require('child_process');
		execSync('rimraf lib/**/*.css lib/**/*.css.map lib/**/*.scss.d.ts', { stdio: 'pipe' });
		console.log('Cleaned existing style files');
	} catch (error) {
		console.log('Clean completed');
	}
	return Promise.resolve();
}

// Watch task
function watchStyles() {
	console.log('Watching SCSS files for changes...');
	let debounceTimeout;

	const watcher = gulp.watch(paths.styles.watch);

	watcher.on('change', (filePath) => {
		console.log(`File changed: ${path.relative(process.cwd(), filePath)}`);

		clearTimeout(debounceTimeout);
		debounceTimeout = setTimeout(() => {
			console.log('Rebuilding styles...');
			gulp.series(
				cleanStyles,
				gulp.parallel(buildRegularStyles, buildModuleStyles, copyScssFiles),
				generateGlobalTypings
			)((err) => {
				if (err) {
					console.error('Build error:', err);
				} else {
					console.log('Styles rebuilt successfully');
				}
			});
		}, 300);
	});

	return watcher;
}

// Main build task
const buildStylesTask = gulp.series(
	cleanStyles,
	gulp.parallel(buildRegularStyles, buildModuleStyles, copyScssFiles),
	generateGlobalTypings
);

// Export tasks
exports.buildStyles = buildStylesTask;
exports['build-styles'] = buildStylesTask;
exports.watchStyles = watchStyles;
exports['watch-styles'] = watchStyles;
exports.cleanStyles = cleanStyles;
exports['clean-styles'] = cleanStyles;
exports.buildRegularStyles = buildRegularStyles;
exports.buildModuleStyles = buildModuleStyles;
exports.copyScssFiles = copyScssFiles;
exports.generateGlobalTypings = generateGlobalTypings;

// Default task
exports.default = buildStylesTask;
