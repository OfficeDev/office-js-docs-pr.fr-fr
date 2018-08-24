---
title: Débogage des compléments Office sur iPad et Mac
description: ''
ms.date: 03/21/2018
ms.openlocfilehash: 24a95799ba400aa6ecf93ee5431cf3fd64779d75
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925205"
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a><span data-ttu-id="a643e-102">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="a643e-102">Debug Office Add-ins on iPad and Mac</span></span>

<span data-ttu-id="a643e-p101">Vous pouvez utiliser Visual Studio pour le développement et le débogage des compléments sur Windows. Toutefois, vous ne pouvez pas l’utiliser pour déboguer les compléments sur iPad ou sur Mac. Dans la mesure où les compléments sont développés dans le code HTML et Javascript, ils devraient fonctionner sur différentes plateformes. Il peut toutefois exister de légères différences dans l’affichage du code HTML dans les différents navigateurs. Cette rubrique explique comment déboguer les compléments en exécution sur iPad ou sur Mac.</span><span class="sxs-lookup"><span data-stu-id="a643e-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac.</span></span> 

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="a643e-106">Débogage avec l'inspecteur Web de Safari sur un Mac</span><span class="sxs-lookup"><span data-stu-id="a643e-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="a643e-107">Si vous avez un complément qui affiche l'interface utilisateur dans un volet de tâches ou dans un complément de contenu, vous pouvez déboguer un complément Office à l'aide de Safari Web Inspector.</span><span class="sxs-lookup"><span data-stu-id="a643e-107">If you have add-in that shows UI in a taskpane or in a content add-in, you can debug an Office add-in using Safari Web Inspector.</span></span> 

<span data-ttu-id="a643e-108">Pour pouvoir déboguer les compléments Office sur Mac, vous devez disposer de Mac OS High Sierra ET de Mac Office Version : 16.9.1 (Build 18012504) ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="a643e-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="a643e-109">Si vous n'avez pas de build Office pour Mac, vous pouvez en obtenir un en rejoignant notre [programme pour les développeurs Office 365](https://aka.ms/o365devprogram).</span><span class="sxs-lookup"><span data-stu-id="a643e-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="a643e-110">Pour commencer, ouvrez un terminal et réglez la propriété `OfficeWebAddinDeveloperExtras` pour l'application Office concernée en procédant comme suit :</span><span class="sxs-lookup"><span data-stu-id="a643e-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="a643e-111">Ensuite, ouvrez l'application Office et insérez votre complément.</span><span class="sxs-lookup"><span data-stu-id="a643e-111">Then, open the Office application and insert your add-in.</span></span> <span data-ttu-id="a643e-112">Cliquez avec le bouton droit sur le complément et vous devriez voir une option **Inspecter l'élément** dans le menu contextuel.</span><span class="sxs-lookup"><span data-stu-id="a643e-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span>  <span data-ttu-id="a643e-113">Sélectionnez cette option et l'inspecteur apparaîtra, où vous pouvez définir des points d'arrêt et déboguer votre complément.</span><span class="sxs-lookup"><span data-stu-id="a643e-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="a643e-114">Veuillez noter qu'il s'agit d'une fonctionnalité expérimentale et il n'y a aucune garantie que nous allons la conserver dans les futures versions des applications Office.</span><span class="sxs-lookup"><span data-stu-id="a643e-114">Please note that this is an experimental feature and there are no guarantees that we will preserve this functionality in future versions of Office applications.</span></span>

## <a name="debugging-with-vorlonjs-on-a-ipad-or-mac"></a><span data-ttu-id="a643e-115">Débogage avec Vorlon.JS sur iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="a643e-115">Debugging with Vorlon.JS on a iPad or Mac</span></span>

<span data-ttu-id="a643e-116">Pour déboguer un complément sur iPad ou Mac, vous pouvez utiliser Vorlon.JS, un débogueur pour pages Web similaire aux outils F12.</span><span class="sxs-lookup"><span data-stu-id="a643e-116">To debug an add-in on iPad or Mac, you can use Vorlon.JS, a debugger for web pages that is similar to the F12 tools.</span></span> <span data-ttu-id="a643e-117">Il est conçu pour fonctionner à distance et déboguer des pages web sur différents appareils.</span><span class="sxs-lookup"><span data-stu-id="a643e-117">It is designed to work remotely and it enables you to debug web pages across different devices.</span></span> <span data-ttu-id="a643e-118">Pour plus d’informations, consultez le [site web de Vorlon](http://www.vorlonjs.com).</span><span class="sxs-lookup"><span data-stu-id="a643e-118">For more information, see the [Vorlon website](http://www.vorlonjs.com).</span></span>  


### <a name="install-and-set-up-vorlonjs"></a><span data-ttu-id="a643e-119">Installation et configuration de Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="a643e-119">Install and set up up Vorlon.JS on a Mac or iPad</span></span>  

1.  <span data-ttu-id="a643e-120">Connectez-vous à l’appareil en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="a643e-120">Log on to the device as an administrator.</span></span>

2.  <span data-ttu-id="a643e-121">Installez [Node.js](https://nodejs.org) s’il n’est pas déjà installé.</span><span class="sxs-lookup"><span data-stu-id="a643e-121">Install [Node.js](https://nodejs.org) if it isn't already installed.</span></span> 

3.  <span data-ttu-id="a643e-p105">Ouvrez une fenêtre **Terminal** et entrez la commande `npm i -g vorlon`. L’outil est installé dans le dossier `/usr/local/lib/node_modules/vorlon`.</span><span class="sxs-lookup"><span data-stu-id="a643e-p105">Open a **Terminal** window and enter the command `npm i -g vorlon`. The tool is installed to `/usr/local/lib/node_modules/vorlon`.</span></span>


### <a name="configure-vorlonjs-to-use-https"></a><span data-ttu-id="a643e-124">Configuration de Vorlon.JS pour une utilisation avec le protocole HTTPS</span><span class="sxs-lookup"><span data-stu-id="a643e-124">Configure Vorlon.JS to use HTTPS</span></span>

<span data-ttu-id="a643e-p106">Pour déboguer une application à l’aide de Vorlon.JS, ajoutez la balise `<script>` à la page d’ouverture de l’application qui charge un script Vorlon.JS à partir d’un emplacement connu (pour plus de détails, reportez-vous à la procédure suivante). Si un complément est sécurisé par SSL (HTTPS), tout script qu’il utilis doit être hébergé sur un serveur HTTPS, y compris le script Vorlon.JS. Par conséquent, afin d’utiliser Vorlon.JS avec des compléments, vous devez le configurer pour qu’il se serve du protocole SSL.</span><span class="sxs-lookup"><span data-stu-id="a643e-p106">To debug an application using Vorlon.JS, you add a `<script>` tag to the opening page of the application that loads a Vorlon.JS script from a well-known location (for details, see the following procedure). If an add-in is SSL-secured (HTTPS), any scripts that it uses must be hosted from an HTTPS server, including the Vorlon.JS script. Therefore, you must configure Vorlon.JS to use SSL in order to use Vorlon.JS with add-ins.</span></span> 

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  <span data-ttu-id="a643e-128">Dans **Finder**, accédez à `/usr/local/lib/node_modules/vorlon`, ouvrez le menu contextuel (en cliquant avec le bouton droit) du dossier `/Server`, puis sélectionnez **Lire les informations**.</span><span class="sxs-lookup"><span data-stu-id="a643e-128">In **Finder**, go to `/usr/local/lib/node_modules/vorlon`, open the context menu for (right-click) the `/Server` folder, and then select **Get Info**.</span></span>

2.  <span data-ttu-id="a643e-129">Cliquez sur l’icône en forme de cadenas dans le coin inférieur droit de la fenêtre **Informations sur le serveur** pour déverrouiller le dossier.</span><span class="sxs-lookup"><span data-stu-id="a643e-129">Choose the padlock icon in the lower right corner of the **Server info** window to unlock the folder.</span></span>

3. <span data-ttu-id="a643e-130">Dans la section **Partage et permissions** de la fenêtre, définissez le **privilège** pour le groupe **personnel** sur **Lecture et écriture**.</span><span class="sxs-lookup"><span data-stu-id="a643e-130">In the **Sharing and Permissions** section of the window, set the **Privilege** for the **staff** group to **Read & Write**.</span></span>

4. <span data-ttu-id="a643e-131">Cliquez à nouveau sur l’icône en forme de cadenas pour ***verrouiller à nouveau*** le dossier.</span><span class="sxs-lookup"><span data-stu-id="a643e-131">Choose the padlock icon again to ***relock*** the folder.</span></span>

5. <span data-ttu-id="a643e-132">Dans **Finder**, développez le sous-dossier `/Server`, cliquez avec le bouton droit sur le fichier `config.json`, puis sélectionnez **Lire les informations**.</span><span class="sxs-lookup"><span data-stu-id="a643e-132">Back in **Finder**, expand the `/Server` subfolder, right-click the file `config.json`, and then select **Get Info**.</span></span>

6. <span data-ttu-id="a643e-p107">Dans la fenêtre **config.json info**, modifiez les privilèges du fichier de la même façon que pour le dossier `/Server` parent. Verrouillez à nouveau le dossier et fermez la fenêtre.</span><span class="sxs-lookup"><span data-stu-id="a643e-p107">In the **config.json info** window, change the privileges of the file exactly the way you did for its parent `/Server` folder. Be sure to relock and close the window.</span></span>

7. <span data-ttu-id="a643e-p108">Dans **Finder**, cliquez avec le bouton droit sur le fichier `config.json`, sélectionnez **Ouvrir avec**, puis **TextEdit**. Le fichier s’ouvre dans un éditeur de texte.</span><span class="sxs-lookup"><span data-stu-id="a643e-p108">Back in **Finder**, right-click the file `config.json`, select **Open with**, and then select **TextEdit**. The file opens in a text editor.</span></span>

8. <span data-ttu-id="a643e-137">Définissez la valeur de la propriété **useSSL** sur `true`.</span><span class="sxs-lookup"><span data-stu-id="a643e-137">Change the value of the **useSSL** property to `true`.</span></span>

9. <span data-ttu-id="a643e-p109">Dans la section **Modules**, recherchez le module ayant pour **ID** `OFFICE` et pour **nom** `Office Addin`. Si la valeur de la propriété **enabled** pour le module n’est pas déjà définie sur `true`, définissez-la sur `true`.</span><span class="sxs-lookup"><span data-stu-id="a643e-p109">In the **plugins** section, find the plugin with the **id** of `OFFICE` and the **name** of `Office Addin`. If the **enabled** property for the plug-in is not already `true`, set it to `true`.</span></span>

10. <span data-ttu-id="a643e-140">Enregistrez le fichier et fermez l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="a643e-140">Save the file and close the editor.</span></span>

11. <span data-ttu-id="a643e-141">Dans **Finder**, accédez à `/usr/local/lib/node_modules/vorlon`, cliquez avec le bouton droit sur le sous-dossier `Server`, et sélectionnez **Nouveau terminal au dossier**.</span><span class="sxs-lookup"><span data-stu-id="a643e-141">In **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
12. <span data-ttu-id="a643e-p110">Dans la fenêtre **Terminal**, entrez `sudo vorlon`. Vous êtes invité à saisir le mot de passe de l’administrateur. Le serveur Vorlon démarre. Laissez la fenêtre **Terminal** ouverte.</span><span class="sxs-lookup"><span data-stu-id="a643e-p110">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

13. <span data-ttu-id="a643e-p111">Ouvrez une fenêtre de navigateur et accédez à `https://localhost:1337`, qui est l’interface de Vorlon.JS. Lorsque vous y êtes invité, sélectionnez **Toujours** pour approuver le certificat de sécurité.</span><span class="sxs-lookup"><span data-stu-id="a643e-p111">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface. When prompted, choose **Always** to trust the security certificate.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="a643e-p112">Si aucune fenêtre d’invite n’apparaît, il se peut que vous deviez approuver le certificat manuellement. Le fichier de certificat est le suivant : `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Suivez la procédure ci-dessous. Si vous rencontrez des problèmes, consultez l’aide de Macintosh ou iPad.</span><span class="sxs-lookup"><span data-stu-id="a643e-p112">If you are not prompted, you might need to trust the certificate manually. The certificate file is `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Try the following steps. If you have trouble, consult Macintosh or iPad help.</span></span> 
    >
    > 1. <span data-ttu-id="a643e-152">Fermez la fenêtre du navigateur et, dans la fenêtre **Terminal** en cours d’exécution sur le serveur Vorlon, utilisez le raccourci Ctrl+C pour arrêter le serveur.</span><span class="sxs-lookup"><span data-stu-id="a643e-152">Close the browser window and in the **Terminal** window that is running the Vorlon server, use Control-C to stop the server.</span></span>
    > 2. <span data-ttu-id="a643e-p113">Dans **Finder**, cliquez avec le bouton droit sur le fichier `server.crt` et sélectionnez **Trousseaux d’accès**. La fenêtre **Trousseaux d’accès** s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="a643e-p113">In **Finder**, right-click the `server.crt` file and select **Keychain Access**. The **Keychain Access** window opens.</span></span>
    > 3. <span data-ttu-id="a643e-p114">Dans la liste **Trousseaux** sur la gauche, sélectionnez **Connexion** si l’option n’est pas déjà sélectionnée, puis choisissez **Certificats** dans la section **Catégorie**. Le certificat **localhost** figure dans la liste.</span><span class="sxs-lookup"><span data-stu-id="a643e-p114">In the **Keychains** list on the left, select **login** if it is not already selected, and then select **Certificates** in the **Category** section. The certificate **localhost** is listed.</span></span>
    > 4. <span data-ttu-id="a643e-p115">Cliquez avec le bouton droit sur le certificat **localhost** et sélectionnez **Lire les informations**. Une fenêtre **localhost** s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="a643e-p115">Right-click the certificate **localhost** and select **Get Info**. A **localhost** window opens.</span></span>
    > 5. <span data-ttu-id="a643e-159">Dans la section **Approuver**, ouvrez le sélecteur nommé **Lors de l’utilisation de ce certificat** et sélectionnez **Toujours approuver**.</span><span class="sxs-lookup"><span data-stu-id="a643e-159">In the **Trust** section, open the selector labeled **When using this certificate** and select **Always Trust**.</span></span> 
    > 6. <span data-ttu-id="a643e-p116">Fermez la fenêtre **localhost**. Si l’action réussit, une croix blanche dans un cercle bleu apparaît sur l’icône du certificat **localhost** dans la fenêtre **Trousseaux d’accès**.</span><span class="sxs-lookup"><span data-stu-id="a643e-p116">Close the **localhost** window. If the action was successful, the **localhost** certificate in the **Keychain Access** window has a white cross in a blue circle on its icon.</span></span>


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a><span data-ttu-id="a643e-162">Configuration du complément pour le débogage Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="a643e-162">Configure the add-in for Vorlon.JS debugging</span></span>

1. <span data-ttu-id="a643e-163">Ajoutez la balise de script suivante à la section `<head>` du fichier home.html (ou fichier HTML principal) de votre complément :</span><span class="sxs-lookup"><span data-stu-id="a643e-163">Add the following script tag to the `<head>` section of the home.html file (or main HTML file) of your add-in:</span></span>

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>    
    ```  

2. <span data-ttu-id="a643e-164">Déployez l’application web du complément sur un serveur web accessible à partir de l’ordinateur Mac ou de l’iPad, tel qu’un site web Azure.</span><span class="sxs-lookup"><span data-stu-id="a643e-164">Deploy the add-in web application to a web server that is accessible from the Mac or iPad, such as an Azure website.</span></span> 

3. <span data-ttu-id="a643e-165">Mettez à jour l’URL du complément à tous les emplacements où elle apparaît dans le manifeste du complément.</span><span class="sxs-lookup"><span data-stu-id="a643e-165">Update the URL of the add-in in all the places where the URL appears in the add-in manifest.</span></span>

4. <span data-ttu-id="a643e-166">Copiez le manifeste du complément dans le dossier suivant sur l’ordinateur Mac ou l’iPad : `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, où *{host_name}* est Word, Excel, PowerPoint ou Outlook.</span><span class="sxs-lookup"><span data-stu-id="a643e-166">Copy the add-in manifest to the following folder on the Mac or iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, where *{host_name}* is Word, Excel, PowerPoint, or Outlook.</span></span>


### <a name="inspect-an-add-in-in-vorlonjs"></a><span data-ttu-id="a643e-167">Vérification d’un complément dans Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="a643e-167">Inspect an add-in in Vorlon.JS</span></span>

1. <span data-ttu-id="a643e-168">Si le serveur Vorlon n’est pas en cours d’exécution, dans **Finder**, accédez à `/usr/local/lib/node_modules/vorlon`, puis cliquez avec le bouton droit sur le sous-dossier `Server` et sélectionnez **Nouveau terminal au dossier**.</span><span class="sxs-lookup"><span data-stu-id="a643e-168">If the Vorlon server is not running, in **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
2.  <span data-ttu-id="a643e-p117">Dans la fenêtre **Terminal**, entrez `sudo vorlon`. Vous êtes invité à saisir le mot de passe de l’administrateur. Le serveur Vorlon démarre. Laissez la fenêtre **Terminal** ouverte.</span><span class="sxs-lookup"><span data-stu-id="a643e-p117">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

3.  <span data-ttu-id="a643e-173">Ouvrez une fenêtre de navigateur et accédez à `https://localhost:1337`, qui est l’interface de Vorlon.JS.</span><span class="sxs-lookup"><span data-stu-id="a643e-173">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface.</span></span>

4. <span data-ttu-id="a643e-p118">Chargez une version test du complément. S’il s’agit d’un complément pour Excel, PowerPoint ou Word, chargez une version test en suivant les étapes décrites dans la rubrique relative au [chargement d’une version test d’un complément Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md). S’il s’agit d’un complément Outlook, chargez une version de test en suivant les étapes décrites dans la rubrique relative au [chargement d’une version test de compléments Outlook à des fins de test](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing). Si le complément n’utilise pas les commandes du complément, il s’ouvre automatiquement. Sinon, cliquez sur le bouton d’ouverture du complément. En fonction de la version de l’application hôte d’Office, vous trouverez le bouton sur l’onglet **Accueil** ou sur l’onglet **Complément**.</span><span class="sxs-lookup"><span data-stu-id="a643e-p118">Sideload the add-in. If it is for Excel, PowerPoint, or Word, sideload it as described in [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md). If it is an Outlook add-in, sideload it as described in [Sideload Outlook Add-ins for testing](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing). If the add-in does not use add-in commands, it will open immediately. Otherwise, choose the button to open the add-in. Depending on the build of the Office host application, the button will be on either the **Home** tab or an **Add-in** tab.</span></span>

<span data-ttu-id="a643e-180">Le complément apparaît dans la liste des clients dans Vorlon.JS (sur la gauche dans l’interface de Vorlon.JS) en tant que **{Système d’exploitation} - n**, pour un nombre *n*, et où *{Système d’exploitation}* correspond au type d’appareil (par exemple, « Macintosh »).</span><span class="sxs-lookup"><span data-stu-id="a643e-180">The add-in will show up in the list of Clients in Vorlon.JS (on the left side of the Vorlon.JS interface) as **{OS} - n**, for some number *n*, and where *{OS}* is the device type, such as "Macintosh".</span></span> 

![Capture d’écran de l’interface Vorlon.js](../images/vorlon-interface.png)

<span data-ttu-id="a643e-p119">L’outil Vorlon dispose d’une variété de plug-ins. Les plug-ins actuellement activés apparaissent sous forme d’onglets dans la partie supérieure de l’interface de l’outil. (Vous pouvez en activer davantage en cliquant sur l’icône en forme d’engrenage sur la gauche.) Ces plug-ins sont semblables aux fonctions disponibles dans les outils F12. Par exemple, vous pouvez mettre en surbrillance les éléments DOM, exécuter des commandes, etc. Pour plus d’informations, reportez-vous à la page relative à la [documentation principale sur les plug-ins Vorlon](http://vorlonjs.com/documentation/#console).</span><span class="sxs-lookup"><span data-stu-id="a643e-p119">The Vorlon tool has a variety of plug-ins. The ones that are currently enabled appear as tabs at the top of the tool. (You can enable more plug-ins by choosing the gears icon on the left.) These plug-ins are  similar to the functions in F12 tools. For example, you can highlight DOM elements, execute commands, and more. For more details, see [Vorlon Documentation Core Plugins](http://vorlonjs.com/documentation/#console)</span></span> 

<span data-ttu-id="a643e-p120">Un plug-in **Complément Office** permet d’ajouter des fonctionnalités supplémentaires pour Office.js, telles que l’exploration du modèle objet, l’exécution d’appels Office.js et la lecture des valeurs des propriétés de l’objet. Pour plus d’informations, reportez-vous à l’article relatif à l’utilisation du [plug-in VorlonJS pour déboguer un complément Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span><span class="sxs-lookup"><span data-stu-id="a643e-p120">An **Office Addin** plug-in adds extra capabilities for Office.js, such as exploring the object model, executing Office.js calls, and reading the values of object properties. For instructions, see [VorlonJS plugin for debugging Office Add-in](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span></span>

> [!NOTE]
> <span data-ttu-id="a643e-188">il n’existe aucun moyen de définir des points d’arrêt dans Vorlon.JS.</span><span class="sxs-lookup"><span data-stu-id="a643e-188">There is no way to set break points in Vorlon.JS.</span></span>


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a><span data-ttu-id="a643e-189">Effacement du cache de l’application Office sur un ordinateur Mac ou un iPad</span><span class="sxs-lookup"><span data-stu-id="a643e-189">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="a643e-p121">Les compléments sont souvent mis en cache dans Office pour Mac, pour des raisons de performances. En règle générale, vous pouvez effacer le cache en rechargeant le complément. En présence de plusieurs compléments dans le même document, il se peut que le processus d’effacement automatique du cache lors du rechargement ne fonctionne pas systématiquement.</span><span class="sxs-lookup"><span data-stu-id="a643e-p121">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span> 

<span data-ttu-id="a643e-193">Sur un ordinateur Mac, vous pouvez effacer le cache manuellement en supprimant tous les éléments contenus dans le dossier `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="a643e-193">On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

<span data-ttu-id="a643e-p122">Sur un iPad, vous pouvez appeler `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.</span><span class="sxs-lookup"><span data-stu-id="a643e-p122">On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>
