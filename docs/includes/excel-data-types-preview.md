> [!NOTE]
> Les API de types de données sont actuellement disponibles uniquement en prévisualisation publique. L’aperçu API peut être modifiés et n’est pas destinés à utiliser dans un environnement de production. Nous vous recommandons de les tester uniquement dans les environnements de test et de développement. N’utilisez pas un aperçu d’API dans un environnement de production ou dans les documents commerciaux importants.
>
> Pour utiliser les API disponibles en préversion :
>
> - Vous devez référencer la **bibliothèque bêta** sur le réseau de distribution de contenu (CDN) ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . Le [fichier de définition de](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) type pour la compilation et la IntelliSense TypeScript se trouve aux CDN et [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Vous pouvez installer ces types avec `npm install --save-dev @types/office-js-preview` . Pour plus d’informations, voir le @microsoft du package NPM [office-js.](https://www.npmjs.com/package/@microsoft/office-js)
> - Vous devrez peut-être rejoindre [Office programme Insider pour](https://insider.office.com) accéder à des builds Office plus récentes.
>
> Pour tester les types de données dans Office sur Windows, vous devez avoir un numéro de build Excel supérieur ou égal à 16.0.14626.10000. Pour tester les types de données dans Office sur Mac, vous devez avoir un numéro de build Excel supérieur ou égal à 16.55.21102600.