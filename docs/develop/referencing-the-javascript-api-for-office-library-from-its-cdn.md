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
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a><span data-ttu-id="eb9ae-102">Référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de distribution de contenu (CDN)</span><span class="sxs-lookup"><span data-stu-id="eb9ae-102">Referencing the JavaScript API for Office library from its content delivery network (CDN)</span></span>

> [!NOTE]
> <span data-ttu-id="eb9ae-103">Outre les étapes décrites dans cet article, si vous souhaitez utiliser TypeScript, puis utiliser Intellisense vous devez exécuterez la commande suivante dans l’invite du système prenant en charge Node (ou la fenêtre git bash) à partir de la racine de votre dossier de projet.</span><span class="sxs-lookup"><span data-stu-id="eb9ae-103">In addition to the steps described in this article, if you want to use TypeScript, then to get Intellisense you will need run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="eb9ae-104">Vous devez avoir [Node.js](https://nodejs.org) installé (qui inclut npm).</span><span class="sxs-lookup"><span data-stu-id="eb9ae-104">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>
> 
> ```
> npm install --save-dev @types/office-js
> ```

<span data-ttu-id="eb9ae-105">La bibliothèque de l’[interface API JavaScript pour Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) comprend le fichier Office.js et des fichiers .js propres aux applications hôtes associées, comme Excel-15.js et Outlook15.js.</span><span class="sxs-lookup"><span data-stu-id="eb9ae-105">The [JavaScript API for Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js.</span></span> 


<span data-ttu-id="eb9ae-106">La façon la plus simple de référencer l’interface API est d’utiliser notre CDN en ajoutant le `<script>` suivant à la balise `<head>` de votre page :</span><span class="sxs-lookup"><span data-stu-id="eb9ae-106">The simplest way to reference the API is to use our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="eb9ae-p102">La valeur `/1/` devant `office.js` dans l’URL du CDN indique la dernière version incrémentielle comprise dans la version 1 d’Office.js. Étant donné que l’interface API JavaScript pour Office maintient la compatibilité descendante, la dernière version continuera de prendre en charge les membres de l’API ajoutés précédemment dans la version 1. Si vous devez mettre à jour un projet existant, consultez la rubrique relative à la [mise à jour de la version de votre interface API JavaScript pour Office et des fichiers de schéma de manifeste](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="eb9ae-p102">The  `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="eb9ae-p103">Si vous envisagez de publier votre complément Office à partir d’AppSource, vous devez utiliser cette référence au CDN. Les références locales sont adaptées uniquement au développement interne et au débogage des scénarios.</span><span class="sxs-lookup"><span data-stu-id="eb9ae-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!IMPORTANT]
>  <span data-ttu-id="eb9ae-p104">Lorsque vous développez un complément pour une application hôte Office, référencez interface API JavaScript pour Office à partir de l’intérieur de la section `<head>` de la page. Ainsi, l’API est entièrement initialisée avant les éléments Body. Les hôtes Office exigent que les compléments soient initialisés 5 secondes après l’activation. Si votre complément n’est pas activé dans ce délai, il sera déclaré comme bloqué et un message d’erreur sera affiché à l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="eb9ae-p104">When you develop an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements. Office hosts require that add-ins initialize within 5 seconds of activation. If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>       

## <a name="see-also"></a><span data-ttu-id="eb9ae-116">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="eb9ae-116">See also</span></span>

- [<span data-ttu-id="eb9ae-117">Présentation de l’interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="eb9ae-117">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="eb9ae-118">Interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="eb9ae-118">JavaScript API for Office</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)
    
