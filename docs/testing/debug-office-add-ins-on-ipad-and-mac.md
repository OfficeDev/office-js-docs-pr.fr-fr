---
title: D?bogage des compl?ments Office sur iPad et Mac
description: ''
ms.date: 03/21/2018
ms.openlocfilehash: 5d68fa000e19d81ebbcd1b383a790958f2bbac72
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a>D?bogage des compl?ments Office sur iPad et Mac

Vous pouvez utiliser Visual Studio pour le d?veloppement et le d?bogage des compl?ments sur Windows. Toutefois, vous ne pouvez pas l?utiliser pour d?boguer les compl?ments sur iPad ou sur Mac. Dans la mesure o? les compl?ments sont d?velopp?s dans le code HTML et Javascript, ils devraient fonctionner sur diff?rentes plateformes. Il peut toutefois exister de l?g?res diff?rences dans l?affichage du code HTML dans les diff?rents navigateurs. Cette rubrique explique comment d?boguer les compl?ments en ex?cution sur iPad ou sur Mac. 

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>D?bogage avec l'inspecteur Web de Safari sur un Mac

Vous pouvez d?boguer un compl?ment Office ? l'aide de l'inspecteur Web de Safari. 

Pour pouvoir d?boguer les compl?ments Office sur Mac, vous devez disposer de Mac OS High Sierra ET de Mac Office Version : 16.9.1 (Build 18012504) ou version ult?rieure. Si vous n'avez pas de build Office pour Mac, vous pouvez en obtenir un en rejoignant notre [programme pour les d?veloppeurs Office 365](https://aka.ms/o365devprogram).

Pour commencer, ouvrez un terminal et r?glez la propri?t? `OfficeWebAddinDeveloperExtras` pour l'application Office concern?e en proc?dant comme suit?:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

Ensuite, ouvrez l'application Office et ins?rez votre compl?ment. Cliquez avec le bouton droit sur le compl?ment et vous devriez voir une option **Inspecter l'?l?ment** dans le menu contextuel.  S?lectionnez cette option et l'inspecteur appara?tra, o? vous pouvez d?finir des points d'arr?t et d?boguer votre compl?ment.

> [!NOTE]
> Veuillez noter qu'il s'agit d'une fonctionnalit? exp?rimentale et il n'y a aucune garantie que nous allons la conserver dans les futures versions des applications Office.

## <a name="debugging-with-vorlonjs-on-a-ipad-or-mac"></a>D?bogage avec Vorlon.JS sur iPad ou Mac

Pour d?boguer un compl?ment sur iPad ou Mac, vous pouvez utiliser Vorlon.JS, un d?bogueur pour pages Web similaire aux outils F12. Il est con?u pour fonctionner ? distance et d?boguer des pages web sur diff?rents appareils. Pour plus d?informations, consultez le [site web de Vorlon](http://www.vorlonjs.com).  


### <a name="install-and-set-up-vorlonjs"></a>Installation et configuration de Vorlon.JS  

1.  Connectez-vous au support en tant qu?administrateur.

2.  Installez [Node.js](https://nodejs.org) s?il n?est pas d?j? install?. 

3.  Ouvrez une fen?tre **Terminal** et entrez la commande `npm i -g vorlon`. L?outil est install? dans le dossier `/usr/local/lib/node_modules/vorlon`.


### <a name="configure-vorlonjs-to-use-https"></a>Configuration de Vorlon.JS pour une utilisation avec le protocole HTTPS

Pour d?boguer une application ? l?aide de Vorlon.JS, ajoutez la balise `<script>` ? la page d?ouverture de l?application qui charge un script Vorlon.JS ? partir d?un emplacement connu (pour plus de d?tails, reportez-vous ? la proc?dure suivante). Si un compl?ment est s?curis? par SSL (HTTPS), tout script qu?il utilis doit ?tre h?berg? sur un serveur HTTPS, y compris le script Vorlon.JS. Par cons?quent, afin d?utiliser Vorlon.JS avec des compl?ments, vous devez le configurer pour qu?il se serve du protocole SSL. 

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  Dans **Finder**, acc?dez ? `/usr/local/lib/node_modules/vorlon`, ouvrez le menu contextuel (en cliquant avec le bouton droit) du dossier `/Server`, puis s?lectionnez **Lire les informations**.

2.  Cliquez sur l?ic?ne en forme de cadenas dans le coin inf?rieur droit de la fen?tre **Informations sur le serveur** pour d?verrouiller le dossier.

3. Dans la section **Partage et permissions** de la fen?tre, d?finissez le **privil?ge** pour le groupe **personnel** sur **Lecture et ?criture**.

4. Cliquez ? nouveau sur l?ic?ne en forme de cadenas pour ***verrouiller ? nouveau*** le dossier.

5. Dans **Finder**, d?veloppez le sous-dossier `/Server`, cliquez avec le bouton droit sur le fichier `config.json`, puis s?lectionnez **Lire les informations**.

6. Dans la fen?tre **config.json info**, modifiez les privil?ges du fichier de la m?me fa?on que pour le dossier `/Server` parent. Verrouillez ? nouveau le dossier et fermez la fen?tre.

7. Dans **Finder**, cliquez avec le bouton droit sur le fichier `config.json`, s?lectionnez **Ouvrir avec**, puis **TextEdit**. Le fichier s?ouvre dans un ?diteur de texte.

8. D?finissez la valeur de la propri?t? **useSSL** sur `true`.

9. Dans la section **Modules**, recherchez le module ayant pour **ID** `OFFICE` et pour **nom** `Office Addin`. Si la valeur de la propri?t? **enabled** pour le module n?est pas d?j? d?finie sur `true`, d?finissez-la sur `true`.

10. Enregistrez le fichier et fermez l??diteur.

11. Dans **Finder**, acc?dez ? `/usr/local/lib/node_modules/vorlon`, cliquez avec le bouton droit sur le sous-dossier `Server`, et s?lectionnez **Nouveau terminal au dossier**. 
    
12. Dans la fen?tre **Terminal**, entrez `sudo vorlon`. Vous ?tes invit? ? saisir le mot de passe de l?administrateur. Le serveur Vorlon d?marre. Laissez la fen?tre **Terminal** ouverte.

13. Ouvrez une fen?tre de navigateur et acc?dez ? `https://localhost:1337`, qui est l?interface de Vorlon.JS. Lorsque vous y ?tes invit?, s?lectionnez **Toujours** pour approuver le certificat de s?curit?. 

    > [!NOTE]
    > Si aucune fen?tre d?invite n?appara?t, il se peut que vous deviez approuver le certificat manuellement. Le fichier de certificat est le suivant : `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Suivez la proc?dure ci-dessous. Si vous rencontrez des probl?mes, consultez l?aide de Macintosh ou iPad. 
    >
    > 1. Fermez la fen?tre du navigateur et, dans la fen?tre **Terminal** en cours d?ex?cution sur le serveur Vorlon, utilisez le raccourci Ctrl+C pour arr?ter le serveur.
    > 2. Dans **Finder**, cliquez avec le bouton droit sur le fichier `server.crt` et s?lectionnez **Trousseaux d?acc?s**. La fen?tre **Trousseaux d?acc?s** s?ouvre.
    > 3. Dans la liste **Trousseaux** sur la gauche, s?lectionnez **Connexion** si l?option n?est pas d?j? s?lectionn?e, puis choisissez **Certificats** dans la section **Cat?gorie**. Le certificat **localhost** figure dans la liste.
    > 4. Cliquez avec le bouton droit sur le certificat **localhost** et s?lectionnez **Lire les informations**. Une fen?tre **localhost** s?ouvre.
    > 5. Dans la section **Approuver**, ouvrez le s?lecteur nomm? **Lors de l?utilisation de ce certificat** et s?lectionnez **Toujours approuver**. 
    > 6. Fermez la fen?tre **localhost**. Si l?action r?ussit, une croix blanche dans un cercle bleu appara?t sur l?ic?ne du certificat **localhost** dans la fen?tre **Trousseaux d?acc?s**.


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a>Configuration du compl?ment pour le d?bogage Vorlon.JS

1. Ajoutez la balise de script suivante ? la section `<head>` du fichier home.html (ou fichier HTML principal) de votre compl?ment :

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>    
    ```  

2. D?ployez l?application web du compl?ment sur un serveur web accessible ? partir de l?ordinateur Mac ou de l?iPad, tel qu?un site web Azure. 

3. Mettez ? jour l?URL du compl?ment ? tous les emplacements o? elle appara?t dans le manifeste du compl?ment.

4. Copiez le manifeste du compl?ment dans le dossier suivant sur l?ordinateur Mac ou l?iPad : `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, o? *{host_name}* est Word, Excel, PowerPoint ou Outlook.


### <a name="inspect-an-add-in-in-vorlonjs"></a>V?rification d?un compl?ment dans Vorlon.JS

1. Si le serveur Vorlon n?est pas en cours d?ex?cution, dans **Finder**, acc?dez ? `/usr/local/lib/node_modules/vorlon`, puis cliquez avec le bouton droit sur le sous-dossier `Server` et s?lectionnez **Nouveau terminal au dossier**. 
    
2.  Dans la fen?tre **Terminal**, entrez `sudo vorlon`. Vous ?tes invit? ? saisir le mot de passe de l?administrateur. Le serveur Vorlon d?marre. Laissez la fen?tre **Terminal** ouverte.

3.  Ouvrez une fen?tre de navigateur et acc?dez ? `https://localhost:1337`, qui est l?interface de Vorlon.JS.

4. Chargez une version test du compl?ment. S?il s?agit d?un compl?ment pour Excel, PowerPoint ou Word, chargez une version test en suivant les ?tapes d?crites dans la rubrique relative au [chargement d?une version test d?un compl?ment Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md). S?il s?agit d?un compl?ment Outlook, chargez une version de test en suivant les ?tapes d?crites dans la rubrique relative au [chargement d?une version test de compl?ments Outlook ? des fins de test](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing). Si le compl?ment n?utilise pas les commandes du compl?ment, il s?ouvre automatiquement. Sinon, cliquez sur le bouton d?ouverture du compl?ment. En fonction de la version de l?application h?te d?Office, vous trouverez le bouton sur l?onglet **Accueil** ou sur l?onglet **Compl?ment**.

Le compl?ment appara?t dans la liste des clients dans Vorlon.JS (sur la gauche dans l?interface de Vorlon.JS) en tant que **{Syst?me d?exploitation} - n**, pour un nombre *n*, et o? *{Syst?me d?exploitation}* correspond au type d?appareil (par exemple, ? Macintosh ?). 

![Capture d??cran de l?interface Vorlon.js](../images/vorlon-interface.png)

L?outil Vorlon dispose d?une vari?t? de plug-ins. Les plug-ins actuellement activ?s apparaissent sous forme d?onglets dans la partie sup?rieure de l?interface de l?outil. (Vous pouvez en activer davantage en cliquant sur l?ic?ne en forme d?engrenage sur la gauche.) Ces plug-ins sont semblables aux fonctions disponibles dans les outils F12. Par exemple, vous pouvez mettre en surbrillance les ?l?ments DOM, ex?cuter des commandes, etc. Pour plus d?informations, reportez-vous ? la page relative ? la [documentation principale sur les plug-ins Vorlon](http://vorlonjs.com/documentation/#console). 

Un plug-in **Compl?ment Office** permet d?ajouter des fonctionnalit?s suppl?mentaires pour Office.js, telles que l?exploration du mod?le objet, l?ex?cution d?appels Office.js et la lecture des valeurs des propri?t?s de l?objet. Pour plus d?informations, reportez-vous ? l?article relatif ? l?utilisation du [plug-in VorlonJS pour d?boguer un compl?ment Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).

> [!NOTE]
> il n?existe aucun moyen de d?finir des points d?arr?t dans Vorlon.JS.


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a>Effacement du cache de l?application Office sur un ordinateur Mac ou un iPad

Les compl?ments sont souvent mis en cache dans Office pour Mac, pour des raisons de performances. En r?gle g?n?rale, vous pouvez effacer le cache en rechargeant le compl?ment. En pr?sence de plusieurs compl?ments dans le m?me document, il se peut que le processus d?effacement automatique du cache lors du rechargement ne fonctionne pas syst?matiquement. 

Sur un ordinateur Mac, vous pouvez effacer le cache manuellement en supprimant tous les ?l?ments contenus dans le dossier `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`. 

Sur un iPad, vous pouvez appeler `window.location.reload(true)` ? partir de JavaScript dans le compl?ment pour forcer le rechargement. Vous pouvez ?galement choisir de r?installer Office.
