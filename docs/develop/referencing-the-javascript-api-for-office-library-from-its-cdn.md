---
title: Référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de distribution de contenu (CDN)
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 422cbd947dde09a8cd19559db9a86ddacd5e2dba
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348092"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de distribution de contenu (CDN)

> [!NOTE]
> Outre les étapes décrites dans cet article, si vous souhaitez utiliser TypeScript, puis utiliser Intellisense vous devez exécuterez la commande suivante dans l’invite du système prenant en charge Node (ou la fenêtre git bash) à partir de la racine de votre dossier de projet. Vous devez avoir [Node.js](https://nodejs.org) installé (qui inclut npm).
> 
> ```
> npm install --save-dev @types/office-js
> ```

La bibliothèque de l’[interface API JavaScript pour Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) comprend le fichier Office.js et des fichiers .js propres aux applications hôtes associées, comme Excel-15.js et Outlook15.js. 


La façon la plus simple de référencer l’interface API est d’utiliser notre CDN en ajoutant le `<script>` suivant à la balise `<head>` de votre page :  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

La valeur `/1/` devant `office.js` dans l’URL du CDN indique la dernière version incrémentielle comprise dans la version 1 d’Office.js. Étant donné que l’interface API JavaScript pour Office maintient la compatibilité descendante, la dernière version continuera de prendre en charge les membres de l’API ajoutés précédemment dans la version 1. Si vous devez mettre à jour un projet existant, consultez la rubrique relative à la [mise à jour de la version de votre interface API JavaScript pour Office et des fichiers de schéma de manifeste](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Si vous envisagez de publier votre complément Office à partir d’AppSource, vous devez utiliser cette référence au CDN. Les références locales sont adaptées uniquement au développement interne et au débogage des scénarios.

> [!IMPORTANT]
>  Lorsque vous développez un complément pour une application hôte Office, référencez interface API JavaScript pour Office à partir de l’intérieur de la section `<head>` de la page. Ainsi, l’API est entièrement initialisée avant les éléments Body. Les hôtes Office exigent que les compléments soient initialisés 5 secondes après l’activation. Si votre complément n’est pas activé dans ce délai, il sera déclaré comme bloqué et un message d’erreur sera affiché à l’utilisateur.       

## <a name="see-also"></a>Voir aussi

- [Présentation de l’interface API JavaScript pour Office](understanding-the-javascript-api-for-office.md)    
- [Interface API JavaScript pour Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)
    
