---
title: Chargement de compléments Office à des fins de test à partir d’un partage réseau
description: Découvrez comment chargement un complément Office à des fins de test à partir d’un partage réseau
ms.date: 06/02/2020
localization_priority: Normal
ms.openlocfilehash: 268fb79c6340aa2d0b8e8278683a0c47b3b60c0e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611245"
---
# <a name="sideload-office-add-ins-for-testing-from-a-network-share"></a>Chargement de compléments Office à des fins de test à partir d’un partage réseau

Vous pouvez tester un complément Office dans un client Office qui se trouve sur Windows en publiant le manifeste sur un partage de fichiers réseau (instructions ci-dessous). Cette option de déploiement est destinée à être utilisée lorsque vous avez terminé le développement et le test sur un hôte local et que vous souhaitez tester le complément à partir d’un serveur non local ou d’un compte Cloud.

> [!IMPORTANT]
> Le déploiement par partage réseau n’est pas pris en charge pour les compléments de production. Cette méthode présente les limitations suivantes :
> 
> - Le complément peut uniquement être installé sur les ordinateurs Windows.
> - Si une nouvelle version d’un complément modifie le ruban, chaque utilisateur doit réinstaller le complément...


> [!NOTE]
> Si votre projet de complément a été créé avec une version suffisamment récente du [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office), le complément se charge automatiquement en version de test dans le client de bureau Office lors de l’exécution de `npm start`.

Cet article s’applique uniquement au test des compléments Word, Excel, PowerPoint et Project et uniquement sur Windows. Si vous souhaitez tester sur une autre plateforme ou tester un complément Outlook, consultez une des rubriques suivantes pour charger une version de votre complément :

- [Chargement de versions test des compléments Office dans Office sur le web](sideload-office-add-ins-for-testing.md)
- [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Chargement de version test des compléments Outlook pour les tester](../outlook/sideload-outlook-add-ins-for-testing.md)

La vidéo suivante présente la procédure de chargement de version test de votre complément dans Office sur le web ou le bureau à l’aide d’un catalogue de dossiers partagés.  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a>Partager un dossier

1. Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.

2. Ouvrez le menu contextuel pour le dossier que vous souhaitez utiliser comme catalogue de dossiers partagés (cliquez sur le dossier avec le bouton droit) et choisissez **Propriétés**.

3. Dans la boîte de dialogue **Propriétés**, ouvrez l’onglet **Partage**, puis choisissez le bouton **Partager**.

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le bouton Partager mis en évidence](../images/sideload-windows-properties-dialog.png)

4. Dans la boîte de dialogue **Accès réseau**, ajoutez-vous ainsi que les autres utilisateurs et/ou groupes avec lesquels vous souhaitez partager votre complément. Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier. Une fois que vous avez choisi les utilisateurs avec lesquels vous souhaitez effectuer le partage, sélectionnez le bouton **Partager**.

5. Lorsqu’un message de confirmation indiquant que **votre dossier est partagé** apparaît, notez le chemin d’accès complet du réseau qui s’affiche juste après le nom du dossier. (Vous devrez entrer cette valeur comme **URL du catalogue** lorsque vous [spécifierez le dossier partagé comme un catalogue approuvé](#specify-the-shared-folder-as-a-trusted-catalog), tel que décrit dans la section suivante de cet article.) Sélectionnez le bouton **Terminé** pour fermer la boîte de dialogue **Accès réseau**.

   ![Boîte de dialogue Accès réseau avec le chemin d’accès partagé mis en évidence](../images/sideload-windows-network-access-dialog.png)

6. Choisissez le bouton **Fermer** pour fermer la boîte de dialogue **Propriétés**.

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>Spécifier le dossier partagé en tant que catalogue approuvé

### <a name="configure-the-trust-manually"></a>Configurer l’approbation manuellement

1. Ouvrez un nouveau document dans Excel, Word, PowerPoint ou Project.

2. Choisissez l’onglet **Fichier**, puis choisissez **Options**.

3. Choisissez l’onglet **Fichier**, puis choisissez **Options**.

4. Choisissez **Catalogues de compléments approuvés**.

5. Dans la zone **URL du catalogue**, entrez le chemin d’accès complet du réseau vers le dossier que vous avez [partagé](#share-a-folder) précédemment. Si vous n’avez pas noté le chemin d’accès complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante.

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le chemin d’accès du réseau mis en évidence](../images/sideload-windows-properties-dialog-2.png)

6. Après avoir entré le chemin d’accès complet du réseau du dossier dans la zone **URL du catalogue**, choisissez le bouton **Ajouter un catalogue**.

7. Cochez la case **Afficher dans le menu** pour l’élément nouvellement ajouté, puis choisissez le bouton **OK** pour fermer la boîte de dialogue **Centre de gestion de la confidentialité**. 

    ![Boîte de dialogue Centre de gestion de la confidentialité avec le catalogue sélectionné](../images/sideload-windows-trust-center-dialog.png)

8. Cliquez sur le bouton **OK** pour fermer la boîte de dialogue **options** .

9. Fermez et ouvrez de nouveau l’application Office afin que vos modifications prennent effet.

### <a name="configure-the-trust-with-a-registry-script"></a>Configurer l’approbation à l’aide d’un script du Registre

1. Dans un éditeur de texte, créez un fichier nommé TrustNetworkShareCatalog.reg.

2. Ajoutez le contenu suivant au fichier :

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```
3. Utilisez l’un des nombreux outils de génération de GUID en ligne, tels que le [Générateur de GUID](https://guidgenerator.com/), pour générer un GUID aléatoire, et dans le fichier TrustNetworkShareCatalog.reg, remplacez la chaîne « -Random-GUID-here- » *dans les deux emplacements* par le GUID. (Les symboles `{}` englobantes doivent subsister).

4. Remplacez la valeur`Url`, par le chemin d’accès complet du réseau vers le dossier que vous avez [partagé](#share-a-folder) précédemment. (Notez que les caractères `\` de l’URL doivent être doublés) Si vous n’avez pas noté le chemin d’accès complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante.

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le chemin d’accès du réseau mis en évidence](../images/sideload-windows-properties-dialog-2.png)

5. Le fichier doit désormais se présenter comme suit. Enregistrez-le.

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

6. Fermez *toutes* les applications Office.

7. Exécutez le fichier TrustNetworkShareCatalog.reg comme vous le feriez pour n’importe quel exécutable, par exemple, double-cliquez sur celui-ci.

## <a name="sideload-your-add-in"></a>Charger une version test de votre complément

1. Placez le fichier XML manifeste d’un complément en cours de test dans le catalogue de dossiers partagés. Notez que vous déployez l’application web sur un serveur web. Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban. Dans Project, sélectionnez **Mes compléments** sous l’onglet **Project** du ruban.

3. Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.

4. Sélectionnez le nom du complément, puis choisissez **OK** pour insérer celui-ci.

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un complément versions test chargées

Vous pouvez supprimer un complément précédemment versions test chargées en effaçant le cache Office sur votre ordinateur. Pour plus d’informations sur la façon d’effacer le cache sur Windows, consultez l’article [effacer le cache Office](clear-cache.md#clear-the-office-cache-on-windows).

## <a name="see-also"></a>Voir aussi

- [Valider le manifeste d’un complément Office](troubleshoot-manifest.md)
- [Vider le cache Office](clear-cache.md)
- [Publier votre complément Office](../publish/publish.md)
