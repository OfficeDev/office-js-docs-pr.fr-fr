---
title: Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office
description: Découvrez comment résoudre les erreurs de développement dans les Office de développement.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2a17a9eafd91cd174209b1974eea61715385c0ad
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990802"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office

Voici une liste des problèmes courants que vous pouvez rencontrer lors du développement d’un Office de développement.

> [!TIP]
> L’effacement Office cache de données résout souvent les problèmes liés au code obsolète. Cela garantit que le dernier manifeste est téléchargé à l’aide des noms de fichiers, du texte du menu et d’autres éléments de commande actuels. Pour plus d’informations, voir [Effacer le cache Office cache.](clear-cache.md)

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Le complément ne se charge pas dans le volet des tâches ou d’autres problèmes existent avec le manifeste du complément

Voir [Valider le manifeste d’un complément Office](troubleshoot-manifest.md) et [Déboguer votre complément avec la journalisation runtime](runtime-logging.md) pour déboguer les problèmes de manifeste de compléments.

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Les modifications apportées aux commandes de complément, y compris les éléments de menu et les boutons du ruban ne s’appliquent pas

L’effacement du cache permet de s’assurer que la dernière version du manifeste de votre add-in est utilisée. Pour effacer le cache Office de données, suivez les instructions de [la Office cache.](clear-cache.md) Si vous utilisez Office sur le Web, effacer le cache de votre navigateur via l’interface utilisateur du navigateur.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Les modifications apportées aux fichiers statiques, tels que JavaScript, HTML et CSS ne sont pas prises en compte.

Le navigateur web met peut-être le contenu de ces fichiers en cache. Pour éviter cela, vous pouvez désactiver la mise en cache côté client lors du développement. Les spécifications dépendent du serveur utilisé. Dans la plupart des cas, elle implique l’ajout d’en-têtes aux réponses HTTP. Nous vous suggérons l’ensemble suivant.

- Cache-Control : « privé, aucun cache, aucun magasin »
- Pragma : « aucun cache »
- Date d’expiration : « -1 »

Un exemple d’opération dans un serveur Node.JS Express est disponible dans [ce fichier app.js](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js). Un exemple de projet ASP.NET est disponible dans [ce fichier cshtml](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>Les modifications apportées aux valeurs des propriétés ne se produisent pas et il n’existe aucun message d’erreur

Consultez la documentation de référence de la propriété pour voir si elle est en lecture seule. En outre, les [définitions TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) pour Office JS spécifient quelles propriétés d’objet sont en lecture seule. Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue silencieusement, sans qu’aucune erreur ne soit lancée. L’exemple suivant tente par erreur de définir la propriété en lecture seule [Chart.id](/javascript/api/excel/excel.chart#id). Voir aussi [Certaines propriétés ne peuvent pas être définies directement.](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>Obtention d’une erreur : « Ce module n’est plus disponible »

Voici quelques-unes des causes de cette erreur. Si vous découvrez d’autres causes, n’hésitez pas à nous en faire part avec l’outil de commentaires en bas de la page.

- Si vous utilisez Visual Studio, il se peut qu’il y a un problème avec le chargement de version secondaire. Fermez toutes les instances de l’hôte Office et des Visual Studio. Redémarrez Visual Studio puis réessayez d’appuyer sur F5.
- Le manifeste du add-in a été supprimé de son emplacement de déploiement, tel qu’un déploiement centralisé, un catalogue SharePoint ou un partage réseau.
- La valeur de [l’élément ID](../reference/manifest/id.md) dans le manifeste a été modifiée directement dans la copie déployée. Si, pour une raison quelconque, vous souhaitez modifier cet ID, supprimez d’abord le module de l’hôte Office, puis remplacez le manifeste d’origine par le manifeste modifié. Vous devez effacer le cache Office pour supprimer toutes les traces de l’original. Consultez [l’article Effacer Office cache pour](clear-cache.md) obtenir des instructions sur l’effacement du cache pour votre système d’exploitation.
- Le manifeste du add-in a un qui n’est pas défini n’importe où dans la section Resources du manifeste, ou il y a une insmatance dans l’orthographe de l’endroit où il est utilisé et où il est défini dans la `resid` [](../reference/manifest/resources.md) `resid` `<Resources>` section.
- Il existe un `resid` attribut quelque part dans le manifeste avec plus de 32 caractères. Un attribut et l’attribut de la ressource correspondante dans la section ne peuvent pas être `resid` `id` plus de `<Resources>` 32 caractères.
- Le add-in possède une commande de add-in personnalisée, mais vous essayez de l’exécuter sur une plateforme qui ne les prend pas en charge. Pour plus d’informations, consultez les ensembles de conditions requises des commandes [de l’autre.](../reference/requirement-sets/add-in-commands-requirement-sets.md)

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>Le add-in ne fonctionne pas sur Edge, mais il fonctionne sur d’autres navigateurs

Voir [Résolution des problèmes Microsoft Edge problèmes.](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel de l’équipe de sécurité envoie des erreurs, mais pas de manière cohérente

Pour [les causes possibles Excel résoudre les](../excel/excel-add-ins-troubleshooting.md) problèmes.

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>Erreurs de validation de schéma de manifeste dans Visual Studio projets

Si vous utilisez des fonctionnalités plus nouvelles qui nécessitent des modifications dans le fichier manifeste, vous pouvez obtenir des erreurs de validation dans Visual Studio. Par exemple, lorsque vous ajoutez l’élément pour implémenter le runtime JavaScript partagé, vous pouvez voir `<Runtimes>` l’erreur de validation suivante.

**L’élément « Host » dans l’espace de noms ' a un élément enfant non valide « Runtimes » dans http://schemas.microsoft.com/office/taskpaneappversionoverrides l’espace de noms http://schemas.microsoft.com/office/taskpaneappversionoverrides ' '**

Si cela se produit, vous pouvez mettre à jour les fichiers XSD Visual Studio aux dernières versions. Les versions de schéma les plus récentes sont à [l'[MS-APPENDIXMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

### <a name="locate-the-xsd-files"></a>Rechercher les fichiers XSD

1. Ouvrez votre projet dans Visual Studio.
1. Dans **l’Explorateur de** solutions, ouvrez manifest.xml fichier. Le manifeste se trouve généralement dans le premier projet sous votre solution.
1. Choose **View**  >  **Properties Window** (F4).
1. Dans la **fenêtre Propriétés,** choisissez les ellipses (...) pour ouvrir l’éditeur de **schémas XML.** Vous trouverez ici l’emplacement exact des dossiers de tous les fichiers de schéma que votre projet utilise.

### <a name="update-the-xsd-files"></a>Mettre à jour les fichiers XSD

1. Ouvrez le fichier XSD que vous souhaitez mettre à jour dans un éditeur de texte. Le nom de schéma de l’erreur de validation correspond au nom de fichier XSD. Par exemple, **ouvrez TaskPaneAppVersionOverridesV1_0.xsd**.
1. Recherchez le schéma mis à jour à [l’emplacement [MS-APPENDIXMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Par exemple, TaskPaneAppVersionOverridesV1_0 est au [schéma taskpaneappversionoverrides](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
1. Copiez le texte dans votre éditeur de texte.
1. Enregistrez le fichier XSD mis à jour.
1. Redémarrez Visual Studio pour récupérer les modifications apportées au nouveau fichier XSD.

Vous pouvez répéter le processus précédent pour tous les schémas supplémentaires qui ne sont pas à jour.

## <a name="see-also"></a>Voir également

- [Débogage de compléments dans Office sur le web](debug-add-ins-in-office-online.md)
- [Charger une version test d’un complément Office sur iPad ou Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Déboguer des compléments Office sur un Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Complément Microsoft Office Extension de débogueur pour Visual Studio Code](debug-with-vs-extension.md)
- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)
- [Déboguer votre complément avec la journalisation runtime](runtime-logging.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](testing-and-troubleshooting.md)
- [Microsoft Q&A (office-js-dev)](/answers/topics/office-js-dev.html)
