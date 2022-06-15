---
title: Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office
description: Découvrez comment résoudre les erreurs de développement dans Office compléments.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7f463b7a7c9a8895283b9f8e18c11bdb63d3da9d
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091124"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office

Voici une liste des problèmes courants que vous pouvez rencontrer lors du développement d’un complément Office.

> [!TIP]
> L’effacement du cache Office résout souvent les problèmes liés au code obsolète. Cela garantit que le dernier manifeste est chargé, à l’aide des noms de fichiers actuels, du texte du menu et d’autres éléments de commande. Pour plus d’informations, consultez [Effacer le cache Office](clear-cache.md).

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Le complément ne se charge pas dans le volet des tâches ou d’autres problèmes existent avec le manifeste du complément

Voir [Valider le manifeste d’un complément Office](troubleshoot-manifest.md) et [Déboguer votre complément avec la journalisation runtime](runtime-logging.md) pour déboguer les problèmes de manifeste de compléments.

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Les modifications apportées aux commandes de complément, y compris les éléments de menu et les boutons du ruban ne s’appliquent pas

L’effacement du cache permet de s’assurer que la dernière version du manifeste de votre complément est utilisée. Pour effacer le cache Office, suivez les instructions fournies dans [Effacer le cache Office](clear-cache.md). Si vous utilisez Office sur le Web, effacez le cache de votre navigateur via l’interface utilisateur du navigateur.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Les modifications apportées aux fichiers statiques, tels que JavaScript, HTML et CSS ne sont pas prises en compte.

Le navigateur web met peut-être le contenu de ces fichiers en cache. Pour éviter cela, vous pouvez désactiver la mise en cache côté client lors du développement. Les spécifications dépendent du serveur utilisé. Dans la plupart des cas, elle implique l’ajout d’en-têtes aux réponses HTTP. Nous vous suggérons l’ensemble suivant.

- Cache-Control : « privé, aucun cache, aucun magasin »
- Pragma : « aucun cache »
- Date d’expiration : « -1 »

Un exemple d’opération dans un serveur Node.JS Express est disponible dans [ce fichier app.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js). Un exemple de projet ASP.NET est disponible dans [ce fichier cshtml](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

Si votre complément est hébergé dans Internet Information Server (IIS), vous pouvez également ajouter ce qui suit à web. config.

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

Si ces étapes ne semblent pas fonctionner au départ, vous devrez peut-être vider le cache du navigateur web. Effectuez cette opération à l’aide de l’interface utilisateur du navigateur web. Il est possible que le cache de périmètre ne soit pas correctement vidé lorsque vous essayez de le faire dans l’interface utilisateur Edge. Si cela se produit, exécutez la commande suivante dans l’invite de commandes Windows.

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>Les modifications apportées aux valeurs de propriété ne se produisent pas et il n’y a aucun message d’erreur

Consultez la documentation de référence de la propriété pour voir si elle est en lecture seule. En outre, les [définitions TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) pour Office JS spécifient les propriétés d’objet en lecture seule. Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue en mode silencieux, sans aucune erreur levée. L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member). Voir aussi [Certaines propriétés ne peuvent pas être définies directement](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>Erreur d’obtention : « Ce complément n’est plus disponible »

Voici quelques-unes des causes de cette erreur. Si vous découvrez des causes supplémentaires, contactez-nous avec l’outil de commentaires en bas de la page.

- Si vous utilisez Visual Studio, il peut y avoir un problème avec le chargement indépendant. Fermez toutes les instances de l’hôte Office et Visual Studio. Redémarrez Visual Studio et réessayez d’appuyer sur F5.
- Le manifeste du complément a été supprimé de son emplacement de déploiement, tel qu’un déploiement centralisé, un catalogue SharePoint ou un partage réseau.
- La valeur de l’élément [ID](/javascript/api/manifest/id) dans le manifeste a été modifiée directement dans la copie déployée. Si, pour une raison quelconque, vous souhaitez modifier cet ID, commencez par supprimer le complément de l’hôte Office, puis remplacez le manifeste d’origine par le manifeste modifié. Vous devez souvent effacer le cache Office pour supprimer toutes les traces de l’original. Consultez [l’article Effacer le cache Office](clear-cache.md) pour obtenir des instructions sur l’effacement du cache pour votre système d’exploitation.
- Le manifeste du complément a une `resid` valeur qui n’est définie nulle part dans la section [Ressources](/javascript/api/manifest/resources) du manifeste, ou il existe une incompatibilité dans l’orthographe de l’emplacement `resid` où il est utilisé et où il est défini dans la `<Resources>` section.
- Il existe un `resid` attribut quelque part dans le manifeste avec plus de 32 caractères. Un `resid` attribut et l’attribut `id` de la ressource correspondante dans la `<Resources>` section ne peuvent pas dépasser 32 caractères.
- Le complément a une commande de complément personnalisée, mais vous essayez de l’exécuter sur une plateforme qui ne les prend pas en charge. Pour plus d’informations, consultez les [ensembles de conditions requises pour les commandes](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets) de complément.

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>Le complément ne fonctionne pas sur Edge, mais il fonctionne sur d’autres navigateurs

Voir [Résolution des problèmes de Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel complément lève des erreurs, mais pas de manière cohérente

Consultez [Résolution des problèmes Excel compléments pour connaître les causes possibles](../excel/excel-add-ins-troubleshooting.md).

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>Erreurs de validation de schéma de manifeste dans Visual Studio projets

Si vous utilisez des fonctionnalités plus récentes qui nécessitent des modifications dans le fichier manifeste, vous pouvez obtenir des erreurs de validation dans Visual Studio. Par exemple, lors de l’ajout de l’élément `<Runtimes>` pour implémenter le runtime JavaScript partagé, vous pouvez voir l’erreur de validation suivante.

**L’élément 'Host' dans l’espace de noms 'http://schemas.microsoft.com/office/taskpaneappversionoverrides' a l’élément enfant 'Runtimes' non valide dans l’espace de noms 'http://schemas.microsoft.com/office/taskpaneappversionoverrides'**

Si cela se produit, vous pouvez mettre à jour les fichiers XSD que Visual Studio utilise aux dernières versions. Les dernières versions du schéma se trouvent dans [[MS-OWEMXML] : Annexe A : Schéma XML complet](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

### <a name="locate-the-xsd-files"></a>Localiser les fichiers XSD

1. Ouvrez votre projet dans Visual Studio.
1. Dans **Explorateur de solutions**, ouvrez le fichier manifest.xml. Le manifeste se trouve généralement dans le premier projet sous votre solution.
1. Choisir **la****fenêtre Propriétés** >  de l’affichage (F4).
1. Dans la **fenêtre Propriétés**, choisissez les points de suspension (...) pour ouvrir l’éditeur **de schémas XML** . Vous trouverez ici l’emplacement exact du dossier de tous les fichiers de schéma utilisés par votre projet.

### <a name="update-the-xsd-files"></a>Mettre à jour les fichiers XSD

1. Ouvrez le fichier XSD à mettre à jour dans un éditeur de texte. Le nom du schéma de l’erreur de validation est corrélé au nom du fichier XSD. Par exemple, ouvrez **TaskPaneAppVersionOverridesV1_0.xsd**.
1. Recherchez le schéma mis à jour à [l’adresse [MS-OWEMXML] : Annexe A : Schéma XML complet](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Par exemple, TaskPaneAppVersionOverridesV1_0 se trouve au [schéma taskpaneappversionoverrides](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
1. Copiez le texte dans votre éditeur de texte.
1. Enregistrez le fichier XSD mis à jour.
1. Redémarrez Visual Studio pour récupérer les nouvelles modifications apportées au fichier XSD.

Vous pouvez répéter le processus précédent pour tous les schémas supplémentaires obsolètes.

## <a name="when-working-offline-no-office-apis-work"></a>Quand vous travaillez hors connexion, aucune API Office ne fonctionne

Lorsque vous chargez la bibliothèque JavaScript Office à partir d’une copie locale au lieu de la CDN, les API peuvent cesser de fonctionner si la bibliothèque n’est pas à jour. Si vous êtes absent d’un projet depuis un certain temps, réinstallez la bibliothèque pour obtenir la dernière version. Le processus varie en fonction de votre IDE. Choisissez l’une des options suivantes en fonction de votre environnement.

- **Visual Studio** : consultez [La mise à jour vers la dernière bibliothèque d’API JavaScript Office](../develop/update-your-javascript-api-for-office-and-manifest-schema-version.md). 
- **Tout autre IDE** : consultez les packages npm [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) et [@types/office-js](https://www.npmjs.com/package/@types/office-js).

## <a name="see-also"></a>Voir également

- [Débogage de compléments dans Office sur le web](debug-add-ins-in-office-online.md)
- [Charger une version test d’un complément Office sur iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Déboguer des compléments Office sur un Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Complément Microsoft Office Extension de débogueur pour Visual Studio Code](debug-with-vs-extension.md)
- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)
- [Déboguer votre complément avec la journalisation runtime](runtime-logging.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md)
- [Microsoft Q&A (office-js-dev)](/answers/topics/office-js-dev.html)
