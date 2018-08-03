---
title: " Chargement de version test de compléments Office"
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: b143999422866dba9b43432359c12f3607261c60
ms.sourcegitcommit: e094aaa06d9aff3d13f8ffd3429d4a31f0b65b81
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/03/2018
ms.locfileid: "21782811"
---
# <a name="sideload-office-add-ins-for-testing"></a> Chargement de version test de compléments Office

Vous pouvez installer un complément Office à tester dans un client Office s’exécutant sous Windows en publiant le manifeste sur un partage de fichiers réseau (instructions ci-dessous).

> [!NOTE]
> Si votre projet de complément a été créé avec l’outil [**Yo Office**](https://github.com/OfficeDev/generator-office), il existe une façon alternative de charger la version test correspondante qui pourrait fonctionner pour vous. Pour plus de détails, voir [Charger une version test des compléments Office à l’aide de la commande de chargement indépendant](sideload-office-addin-using-sideload-command.md).

Cet article s’applique uniquement aux tests des compléments Word, Excel ou PowerPoint sur Windows. Si vous souhaitez tester sur une autre plateforme ou si vous souhaitez tester un complément Outlook, consultez l'une des rubriques suivantes pour charger la version test de votre complément :

- [Chargement de version test des compléments Office dans Office Online](sideload-office-add-ins-for-testing.md)
- [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Chargement de version test de compléments Outlook](../../../../outlook/add-ins/sideload-outlook-add-ins-for-testing)


La vidéo suivante présente vous guide à travers la procédure de chargement indépendant de votre complément dans la version de bureau Office ou Office Online à l’aide du catalogue d'un dossier partagé.  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a>Partager un dossier

1. Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.

2. Ouvrez le menu contextuel du dossier (clic droit), puis choisissez **Propriétés**.

3. Ouvrez l’onglet **Partage**.

4. Dans la page **Choisir les utilisateurs...**, ajoutez votre nom et celui des utilisateurs avec lesquels vous souhaitez partager votre complément. S’ils sont tous membres d’un groupe de sécurité, vous pouvez ajouter le groupe. Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier. 

5. Choisissez **Partager** > **Terminer** > **Fermer**.


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>Spécifier le dossier partagé en tant que catalogue approuvé
      
1. Ouvrez un nouveau document dans Excel, Word ou PowerPoint.
    
2. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
    
3. Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.
    
4. Choisissez **Catalogues de compléments approuvés**.
    
5. Dans la zone **URL du catalogue**, entrez le chemin d’accès réseau complet au catalogue de dossiers partagés, puis choisissez **Ajouter un catalogue**.
    
6. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.

7. Fermez l’application Office afin que vos modifications prennent effet.
    

## <a name="sideload-your-add-in"></a>Charger votre complément

1. Placez le fichier manifeste d’un complément en cours de test dans le catalogue de dossiers partagés. Notez que vous déployez l’application web sur un serveur web. Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.

3. Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.

4. Sélectionnez le nom du complément, puis choisissez **OK** pour insérer le complément.


## <a name="see-also"></a>Voir aussi

- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
- [Publier votre complément Office](../publish/publish.md)
    
