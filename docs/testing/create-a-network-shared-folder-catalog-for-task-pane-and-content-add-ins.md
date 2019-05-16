---
title: Chargement de compléments Office pour des tests
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 24c7719969ddc59d8bb6e525af804515331a51ad
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619043"
---
# <a name="sideload-office-add-ins-for-testing"></a>Chargement de compléments Office pour des tests

Vous pouvez installer un complément Office à des fins de test dans un client Office s’exécutant sur Windows à l’aide d’un catalogue de dossiers partagés pour publier le manifeste sur un partage de fichiers réseau.

> [!NOTE]
> Si votre projet de complément a été créé avec le [générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office), il existe une autre méthode de chargement indépendant pouvant vous convenir. Pour plus de détails, reportez-vous à [Chargement de versions test de compléments Office à l’aide de la commande sideload](sideload-office-addin-using-sideload-command.md).

Cet article s’applique uniquement aux tests de compléments Word, Excel, PowerPoint ou Project sur Windows. Si vous souhaitez tester sur une autre plateforme ou tester un complément Outlook, consultez une des rubriques suivantes pour charger une version de votre complément :

- [Chargement de version test des compléments Office dans Office Online pour les tester](sideload-office-add-ins-for-testing.md)
- [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Chargement de version test des compléments Outlook pour les tester](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

La vidéo suivante présente la procédure de chargement de version test de votre complément dans la version de bureau Office ou Office Online à l’aide d’un catalogue de dossiers partagés.  

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
      
1. Ouvrez un nouveau document dans Excel, Word, PowerPoint ou Project.
    
2. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
    
3. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
    
4. Choisissez **Catalogues de compléments approuvés**.
    
5. Dans la zone **URL du catalogue**, entrez le chemin d’accès complet du réseau vers le dossier que vous avez [partagé](#share-a-folder) précédemment. Si vous n’avez pas noté le chemin d’accès complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante. 

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le chemin d’accès du réseau mis en évidence](../images/sideload-windows-properties-dialog-2.png)
    
6. Après avoir entré le chemin d’accès complet du réseau du dossier dans la zone **URL du catalogue**, choisissez le bouton **Ajouter un catalogue**.

7. Cochez la case **Afficher dans le menu** pour l’élément nouvellement ajouté, puis choisissez le bouton **OK** pour fermer la boîte de dialogue **Centre de gestion de la confidentialité**. 

    ![Boîte de dialogue Centre de gestion de la confidentialité avec le catalogue sélectionné](../images/sideload-windows-trust-center-dialog.png)

8. Sélectionnez le bouton **OK** pour fermer la boîte de dialogue **Options Word**.

9. Fermez et ouvrez de nouveau l’application Office afin que vos modifications prennent effet.
    

## <a name="sideload-your-add-in"></a>Charger une version test de votre complément


1. Placez le fichier XML manifeste d’un complément en cours de test dans le catalogue de dossiers partagés. Notez que vous déployez l’application web sur un serveur web. Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban. Dans Project, sélectionnez **Mes compléments** sous l’onglet **Project** du ruban. 

3. Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.

4. Sélectionnez le nom du complément, puis choisissez **OK** pour insérer le complément.

## <a name="see-also"></a>Voir aussi

- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
- [Publier votre complément Office](../publish/publish.md)
    
