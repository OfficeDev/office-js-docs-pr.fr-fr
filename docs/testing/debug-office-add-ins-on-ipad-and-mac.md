---
title: Débogage des compléments Office sur iPad et Mac
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 5bf626c4c18bcedccd331570b6b892a8c6a903fd
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870400"
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a>Débogage des compléments Office sur iPad et Mac

Vous pouvez utiliser Visual Studio pour le développement et le débogage des compléments sur Windows. Toutefois, vous ne pouvez pas l’utiliser pour déboguer les compléments sur iPad ou sur Mac. Dans la mesure où les compléments sont développés dans le code HTML et Javascript, ils devraient fonctionner sur différentes plateformes. Il peut toutefois exister de légères différences dans l’affichage du code HTML dans les différents navigateurs. Cette rubrique explique comment déboguer les compléments en exécution sur iPad ou sur Mac.

## <a name="debugging-with-vorlonjs-on-ipad-or-mac"></a>Débogage avec Vorlon.JS sur un iPad ou un Mac

Pour déboguer un complément sur iPad ou Mac, vous pouvez utiliser Vorlon.JS, un débogueur pour pages web ressemblant aux outils F12. Il est conçu pour fonctionner à distance et déboguer des pages web sur différents appareils. Pour plus d’informations, consultez le [site web de Vorlon](http://www.vorlonjs.com).  


### <a name="install-and-set-up-vorlonjs"></a>Installer et configurer Vorlon.js  

1.  Connectez-vous à l’appareil en tant qu’administrateur.

2.  Installez [Node.js](https://nodejs.org) s’il n’est pas déjà installé.

3.  Ouvrez une fenêtre **Terminal** et entrez la commande `npm i -g vorlon`. L’outil est installé dans le dossier `/usr/local/lib/node_modules/vorlon`.


### <a name="configure-vorlonjs-to-use-https"></a>Configuration de Vorlon.JS pour une utilisation avec le protocole HTTPS

Pour déboguer une application à l’aide de Vorlon.JS, ajoutez la balise `<script>` à la page d’ouverture de l’application qui charge un script Vorlon.JS à partir d’un emplacement connu (pour plus de détails, reportez-vous à la procédure suivante). Si un complément est sécurisé par SSL (HTTPS), tout script qu’il utilis doit être hébergé sur un serveur HTTPS, y compris le script Vorlon.JS. Par conséquent, afin d’utiliser Vorlon.JS avec des compléments, vous devez le configurer pour qu’il se serve du protocole SSL.

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  Dans **Finder**, accédez à `/usr/local/lib/node_modules/vorlon`, ouvrez le menu contextuel (en cliquant avec le bouton droit) du dossier `/Server`, puis sélectionnez **Lire les informations**.

2.  Cliquez sur l’icône en forme de cadenas dans le coin inférieur droit de la fenêtre **Informations sur le serveur** pour déverrouiller le dossier.

3. Dans la section **Partage et permissions** de la fenêtre, définissez le **privilège** pour le groupe **personnel** sur **Lecture et écriture**.

4. Cliquez à nouveau sur l’icône en forme de cadenas pour ***verrouiller à nouveau*** le dossier.

5. Dans **Finder**, développez le sous-dossier `/Server`, cliquez avec le bouton droit sur le fichier `config.json`, puis sélectionnez **Lire les informations**.

6. Dans la fenêtre **config.json info**, modifiez les privilèges du fichier de la même façon que pour le dossier `/Server` parent. Verrouillez à nouveau le dossier et fermez la fenêtre.

7. Dans **Finder**, cliquez avec le bouton droit sur le fichier `config.json`, sélectionnez **Ouvrir avec**, puis **TextEdit**. Le fichier s’ouvre dans un éditeur de texte.

8. Définissez la valeur de la propriété **useSSL** sur `true`.

9. Dans la section **Modules**, recherchez le module ayant pour **ID** `OFFICE` et pour **nom** `Office Addin`. Si la valeur de la propriété **enabled** pour le module n’est pas déjà définie sur `true`, définissez-la sur `true`.

10. Enregistrez le fichier et fermez l’éditeur.

11. Dans **Finder**, accédez à `/usr/local/lib/node_modules/vorlon`, cliquez avec le bouton droit sur le sous-dossier `Server`, et sélectionnez **Nouveau terminal au dossier**.

12. Dans la fenêtre **Terminal**, entrez `sudo vorlon`. Vous êtes invité à saisir le mot de passe de l’administrateur. Le serveur Vorlon démarre. Laissez la fenêtre **Terminal** ouverte.

13. Ouvrez une fenêtre de navigateur et accédez à `https://localhost:1337`, qui est l’interface de Vorlon.JS. Lorsque vous y êtes invité, sélectionnez **Toujours** pour approuver le certificat de sécurité.

    > [!NOTE]
    > Si aucune fenêtre d’invite n’apparaît, il se peut que vous deviez approuver le certificat manuellement. Le fichier de certificat est le suivant : `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Suivez la procédure ci-dessous. Si vous rencontrez des problèmes, consultez l’aide de Macintosh ou iPad.
    >
    > 1. Fermez la fenêtre du navigateur et, dans la fenêtre **Terminal** en cours d’exécution sur le serveur Vorlon, utilisez le raccourci Ctrl+C pour arrêter le serveur.
    > 2. Dans **Finder**, cliquez avec le bouton droit sur le fichier `server.crt` et sélectionnez **Trousseaux d’accès**. La fenêtre **Trousseaux d’accès** s’ouvre.
    > 3. Dans la liste **Trousseaux** sur la gauche, sélectionnez **Connexion** si l’option n’est pas déjà sélectionnée, puis choisissez **Certificats** dans la section **Catégorie**. Le certificat **localhost** figure dans la liste.
    > 4. Cliquez avec le bouton droit sur le certificat **localhost** et sélectionnez **Lire les informations**. Une fenêtre **localhost** s’ouvre.
    > 5. Dans la section **Approuver**, ouvrez le sélecteur nommé **Lors de l’utilisation de ce certificat** et sélectionnez **Toujours approuver**. 
    > 6. Fermez la fenêtre **localhost**. Si l’action réussit, une croix blanche dans un cercle bleu apparaît sur l’icône du certificat **localhost** dans la fenêtre **Trousseaux d’accès**.


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a>Configuration du complément pour le débogage Vorlon.JS

1. Ajoutez la balise de script suivante à la section `<head>` du fichier home.html (ou fichier HTML principal) de votre complément :

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>
    ```  

2. Déployez l’application web du complément sur un serveur web accessible à partir de l’ordinateur Mac ou de l’iPad, tel qu’un site web Azure.

3. Mettez à jour l’URL du complément à tous les emplacements où elle apparaît dans le manifeste du complément.

4. Copiez le manifeste du complément dans le dossier suivant sur l’ordinateur Mac ou l’iPad : `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, où *{host_name}* est Word, Excel, PowerPoint ou Outlook.


### <a name="inspect-an-add-in-in-vorlonjs"></a>Vérification d’un complément dans Vorlon.JS

1. Si le serveur Vorlon n’est pas en cours d’exécution, dans **Finder**, accédez à `/usr/local/lib/node_modules/vorlon`, puis cliquez avec le bouton droit sur le sous-dossier `Server` et sélectionnez **Nouveau terminal au dossier**. 

2.  Dans la fenêtre **Terminal**, entrez `sudo vorlon`. Vous êtes invité à saisir le mot de passe de l’administrateur. Le serveur Vorlon démarre. Laissez la fenêtre **Terminal** ouverte.

3.  Ouvrez une fenêtre de navigateur et accédez à `https://localhost:1337`, qui est l’interface de Vorlon.JS.

4. Chargez une version test du complément. S’il s’agit d’un complément pour Excel, PowerPoint ou Word, chargez une version test en suivant les étapes décrites dans la rubrique relative au [chargement d’une version test d’un complément Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md). S’il s’agit d’un complément Outlook, chargez une version de test en suivant les étapes décrites dans la rubrique relative au [chargement de version test des compléments Outlook](/outlook/add-ins/sideload-outlook-add-ins-for-testing). Si le complément n’utilise pas les commandes du complément, il s’ouvre automatiquement. Sinon, cliquez sur le bouton d’ouverture du complément. En fonction de la version de l’application hôte d’Office, vous trouverez le bouton sur l’onglet **Accueil** ou sur l’onglet **Complément**.

Le complément apparaît dans la liste des clients dans Vorlon.JS (sur la gauche dans l’interface de Vorlon.JS) en tant que **{Système d’exploitation} - n**, pour un nombre *n*, et où *{Système d’exploitation}* correspond au type d’appareil (par exemple, « Macintosh »).

![Capture d’écran de l’interface Vorlon.js](../images/vorlon-interface.png)

L’outil Vorlon intègre plusieur plug-ins. Ceux qui sont actuellement activés apparaissent sous forme d’onglets dans la partie supérieure de l’interface de l’outil. (Vous pouvez activer d’autres plug-ins en sélectionnant l’icône d’engrenage sur la gauche.) Ces plug-ins sont similaires aux fonctions des outils F12. Par exemple, vous pouvez mettre en surbrillance les éléments DOM, exécuter des commandes, etc. Pour plus d’informations, voir la [documentation principale sur les plug-ins Vorlon](http://vorlonjs.com/documentation/#console).

Un plug-in **Complément Office** permet d’ajouter des fonctionnalités supplémentaires pour Office.js, telles que l’exploration du modèle objet, l’exécution d’appels Office.js et la lecture des valeurs des propriétés de l’objet. Pour plus d’informations, reportez-vous à l’article relatif à l’utilisation du [plug-in VorlonJS pour déboguer un complément Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).

> [!NOTE]
> Il n’existe aucun moyen de définir des points d’arrêt dans Vorlon.JS.

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>Débogage avec l’inspecteur web Safari sur Mac

> [!IMPORTANT]
> Notez que la fonctionnalité**Inspecter l’Élément** est expérimentale et qu’il n’existe aucune garantie que nous conserverons cette fonctionnalité dans les versions futures des applications Office.

Si votre complément affiche une interface utilisateur dans un volet des tâches ou dans un complément de contenu, vous pouvez déboguer un complément Office à l’aide de avec l’inspecteur web Safari.

Pour pouvoir déboguer des compléments Office sur Mac, vous devez disposer de Mac OS High Sierra ET de Mac Office version 16.9.1 (build 18012504) ou version ultérieure. Si vous n’avez pas de build Office pour Mac, vous pouvez en obtenir une en rejoignant le [programme pour les développeurs Office 365](https://aka.ms/o365devprogram).

Pour commencer, ouvrez un terminal, puis définissez la propriété `OfficeWebAddinDeveloperExtras` pour l’application Office pertinente comme suit :

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

Ensuite, ouvrez l’application Office et[insérez votre complément](sideload-an-office-add-in-on-ipad-and-mac.md). Cliquez sur le complément. Vous devriez voir l’option **Inspecter l’élément** s’afficher dans le menu contextuel.  Sélectionnez cette option pour afficher l’inspecteur dans lequel vous pouvez définir des points d’arrêt et déboguer votre complément.

> [!NOTE]
> Si vous essayez d’utiliser l’inspecteur et que la boîte de dialogue scintille, essayez la solution de contournement suivante :
> 1. Pour réduire la taille de la boîte de dialogue.
> 2. Sélectionnez l’option **Inspecter l’élément** qui ouvre une nouvelle fenêtre.
> 3. Redimensionner la boîte de dialogue à sa taille d’origine.
> 4. Utiliser l’inspecteur comme requis.


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a>Effacement du cache de l’application Office sur un ordinateur Mac ou un iPad

Les compléments sont souvent mis en cache dans Office pour Mac, pour des raisons de performances. En règle générale, vous pouvez effacer le cache en rechargeant le complément. En présence de plusieurs compléments dans le même document, il se peut que le processus d’effacement automatique du cache lors du rechargement ne fonctionne pas systématiquement.

Sur un ordinateur Mac, vous pouvez effacer le cache manuellement en supprimant tous les éléments contenus dans le dossier `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

Sur un iPad, vous pouvez appeler `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.
