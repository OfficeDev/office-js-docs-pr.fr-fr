Pour [certaines versions d’Office et de Windows](../concepts/browsers-used-by-office-web-add-ins.md), le moteur JavaScript dans lequel les compléments sont exécutés est fourni par Internet Explorer. Le moteur Internet Explorer ne prend pas en charge les versions de JavaScript ultérieures à ES5. Cela signifie qu’il n’y a pas de gestion spéciale, les fichiers JavaScript que votre complément dessert ne peuvent pas utiliser la syntaxe, les types ou les méthodes qui ont été ajoutés à la langue après ES5. Cela ne signifie pas que vous devez *écrire* dans la syntaxe ES5. Vous disposez de deux autres options :

- Écrivez votre code dans [ECMAScript 2015](https://www.w3schools.com/Js/js_es6.asp) (également appelé ES6) ou JavaScript ultérieur, ou dans une écriture à écrire, puis compilez votre code en ES5 JavaScript à l’aide d’un compilateur tel que [Babel](https://babeljs.io/) ou [TSC](https://www.typescriptlang.org/index.html).
- Écrivez dans ECMAScript 2015 ou une version ultérieure JavaScript, mais chargez également une bibliothèque de [Polyfill](https://wikipedia.org/wiki/Polyfill_(programming)) comme [Core-js](https://github.com/zloirock/core-js) qui permet à Internet Explorer d’exécuter votre code.