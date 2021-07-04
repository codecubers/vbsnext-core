# VbsNext (Core scripts)

Inspired by the revolutionary transformation of Javascript from browser to nodejs and beyond, this project is an attempt to see how far our favorite VB script (.vbs / .wsf) can be extended to give the wings it needed to fly. 

## Usage

```shell
npm install @vbsnext/vbsnext-core
```

Run your primary vbs script through vbsnext CLI to resolve Include(), Echox, Class extends and other keywords

```shell
npx vbsnext <yourscirpt.vbs>
```

Resultant build file will be stored to build direcotry with suffix "-build.vbs"


Suggestions are welcome

### TO-DO
1. Import core libraries as Node packges instead of pre-importing
2. Rendering both vbs and js (Jscirpt)
3. Minify the build vbs scripts
4. Prettify the project vbs scripts
5. Test Assertions Object
6. Test Suite to run all tests