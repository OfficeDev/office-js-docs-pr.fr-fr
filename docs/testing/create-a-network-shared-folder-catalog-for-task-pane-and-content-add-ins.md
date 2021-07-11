---
title: Chargement de Office pour les tester à partir d’un partage réseau
description: Découvrez comment recharger une version test d Office pour le test à partir d’un partage réseau
ms.date: 06/02/2020
localization_priority: Normal
ms.openlocfilehash: 9a44c14669bf0a8fa842e931fc1b12601f73043b
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348306"
---
# <a name="sideload-office-add-ins-for-testing-from-a-network-share"></a>Chargement de Office pour les tester à partir d’un partage réseau

Vous pouvez tester un Office dans un client Office qui se trouve sur Windows en publiant le manifeste sur un partage de fichiers réseau (instructions ci-dessous). Cette option de déploiement est destinée à être utilisée lorsque vous avez terminé le développement et le test sur un localhost et que vous souhaitez tester le module à partir d’un serveur non local ou d’un compte cloud.

> [!IMPORTANT]
> Le déploiement par partage réseau n’est pas pris en charge pour les modules de production. Cette méthode présente les limitations suivantes.
>
> - Le module peut uniquement être installé sur Windows ordinateurs.
> - Si une nouvelle version d’un add-in change le ruban, chaque utilisateur devra réinstaller le module.

> [!NOTE]
> Si votre projet de complément a été créé avec une version suffisamment récente du [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office), le complément se charge automatiquement en version de test dans le client de bureau Office lors de l’exécution de `npm start`.

Cet article s’applique uniquement aux tests de Excel, de PowerPoint et de Project et uniquement sur Windows. Si vous souhaitez tester sur une autre plateforme ou tester un Outlook, consultez l’une des rubriques suivantes pour le chargement de version test de votre module.

- [Chargement de versions test des compléments Office dans Office sur le web](sideload-office-add-ins-for-testing.md)
- [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Chargement de version test des compléments Outlook pour les tester](../outlook/sideload-outlook-add-ins-for-testing.md)

La vidéo suivante présente la procédure de chargement de version test de votre complément dans Office sur le web ou le bureau à l’aide d’un catalogue de dossiers partagés.  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a>Partager un dossier

1. Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.

1. Ouvrez le menu contextuel pour le dossier que vous souhaitez utiliser comme catalogue de dossiers partagés (cliquez sur le dossier avec le bouton droit) et choisissez **Propriétés**.

1. Dans la boîte de dialogue **Propriétés**, ouvrez l’onglet **Partage**, puis choisissez le bouton **Partager**.

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le bouton Partager mis en évidence.](../images/sideload-windows-properties-dialog.png)

1. Dans la boîte de dialogue **Accès réseau**, ajoutez-vous ainsi que les autres utilisateurs et/ou groupes avec lesquels vous souhaitez partager votre complément. Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier. Une fois que vous avez choisi les utilisateurs avec lesquels vous souhaitez effectuer le partage, sélectionnez le bouton **Partager**.

1. Lorsqu’un message de confirmation indiquant que **votre dossier est partagé** apparaît, notez le chemin d’accès complet du réseau qui s’affiche juste après le nom du dossier. (Vous devrez entrer cette valeur comme **URL du catalogue** lorsque vous [spécifierez le dossier partagé comme un catalogue approuvé](#specify-the-shared-folder-as-a-trusted-catalog), tel que décrit dans la section suivante de cet article.) Sélectionnez le bouton **Terminé** pour fermer la boîte de dialogue **Accès réseau**.

   ![Boîte de dialogue Accès réseau avec le chemin d’accès du partage mis en évidence.](../images/sideload-windows-network-access-dialog.png)

1. Choisissez le bouton **Fermer** pour fermer la boîte de dialogue **Propriétés**.

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>Spécifier le dossier partagé en tant que catalogue approuvé

### <a name="configure-the-trust-manually"></a>Configurer l’approbation manuellement

1. Ouvrez un nouveau document dans Excel, Word, PowerPoint ou Project.

1. Choisissez l’onglet **Fichier**, puis choisissez **Options**.

1. Choisissez l’onglet **Fichier**, puis choisissez **Options**.

1. Choisissez **Catalogues de compléments approuvés**.

1. Dans la zone **URL du catalogue**, entrez le chemin d’accès complet du réseau vers le dossier que vous avez [partagé](#share-a-folder) précédemment. Si vous n’avez pas noté le chemin d’accès complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante.

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le chemin d’accès réseau mis en évidence.](../images/sideload-windows-properties-dialog-2.png)

1. Après avoir entré le chemin d’accès complet du réseau du dossier dans la zone **URL du catalogue**, choisissez le bouton **Ajouter un catalogue**.

1. Cochez la case **Afficher dans le menu** pour l’élément nouvellement ajouté, puis choisissez le bouton **OK** pour fermer la boîte de dialogue **Centre de gestion de la confidentialité**. 

    ![Boîte de dialogue Centre de confiance avec le catalogue sélectionné.](../images/sideload-windows-trust-center-dialog.png)

1. Sélectionnez le **bouton OK** pour fermer la boîte de dialogue **Options.**

1. Fermez et ouvrez de nouveau l’application Office afin que vos modifications prennent effet.

### <a name="configure-the-trust-with-a-registry-script"></a>Configurer l’approbation à l’aide d’un script du Registre

1. Dans un éditeur de texte, créez un fichier nommé TrustNetworkShareCatalog.reg.

1. Ajoutez le contenu suivant au fichier.

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```

1. Utilisez l’un des nombreux outils de génération de GUID en ligne, tels que le [Générateur de GUID](https://guidgenerator.com/), pour générer un GUID aléatoire, et dans le fichier TrustNetworkShareCatalog.reg, remplacez la chaîne « -Random-GUID-here- » *dans les deux emplacements* par le GUID. (Les symboles `{}` englobantes doivent subsister).

1. Remplacez la valeur`Url`, par le chemin d’accès complet du réseau vers le dossier que vous avez [partagé](#share-a-folder) précédemment. (Notez que les caractères `\` de l’URL doivent être doublés) Si vous n’avez pas noté le chemin d’accès complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante.

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le chemin d’accès réseau mis en évidence.](../images/sideload-windows-properties-dialog-2.png)

1. Le fichier doit désormais se présenter comme suit. Enregistrez-le.

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

1. Fermez *toutes* les applications Office.

1. Exécutez le fichier TrustNetworkShareCatalog.reg comme vous le feriez pour n’importe quel exécutable, par exemple, double-cliquez sur celui-ci.

## <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. Placez le fichier XML manifeste d’un complément en cours de test dans le catalogue de dossiers partagés. Notez que vous déployez l’application web sur un serveur web. Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

    > [!NOTE]
    > Pour Visual Studio projets, utilisez le manifeste créé par le projet dans le `{projectfolder}\bin\Debug\OfficeAppManifests` dossier.

1. Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban. Dans Project, sélectionnez **Mes compléments** sous l’onglet **Project** du ruban.

1. Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.

1. Sélectionnez le nom du complément, puis choisissez **OK** pour insérer celui-ci.

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un add-in chargé de nouveau

Vous pouvez supprimer un add-in précédemment chargé de nouveau en effantant le cache Office sur votre ordinateur. Vous pouvez trouver plus d’informations sur la façon de Windows cache dans l’article Effacer [le cache Office cache.](clear-cache.md#clear-the-office-cache-on-windows)

## <a name="see-also"></a>Voir aussi

- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)
- [Vider le cache Office](clear-cache.md)
- [Publier votre complément Office](../publish/publish.md)
