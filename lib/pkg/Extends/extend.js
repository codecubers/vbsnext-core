var fs = require("fs")
let lines = fs.readFileSync('./Extends.vbs').toString()
lines = lines.replace(/:/gm, "\r\n")
lines = lines.replace(/\s\s*$/gm, "")
lines = lines.replace(/^\s\s*/gm, "")
lines = lines.replace(/^'(.*)$|'(.*)$/gm, "")
lines = lines.replace(/^'(.*)$/gm, "")
lines = lines.replace(/^Public\s|Private\s/gim, "")
lines = lines.replace(/^(?=\n)$|^\s*|\s*$|\n\n+/gm, "")
// lines = lines.replace(/\s+/g, " ")
// let words = lines.replace(/\s/g, "\r\n")
console.log(lines)
