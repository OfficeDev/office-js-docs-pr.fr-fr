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
# <a name="debug-office-add-ins-on-ipad-and-mac"></a><span data-ttu-id="82bad-102">D?bogage des compl?ments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="82bad-102">Debug Office Add-ins on iPad and Mac</span></span>

<span data-ttu-id="82bad-p101">Vous pouvez utiliser Visual Studio pour le d?veloppement et le d?bogage des compl?ments sur Windows. Toutefois, vous ne pouvez pas l?utiliser pour d?boguer les compl?ments sur iPad ou sur Mac. Dans la mesure o? les compl?ments sont d?velopp?s dans le code HTML et Javascript, ils devraient fonctionner sur diff?rentes plateformes. Il peut toutefois exister de l?g?res diff?rences dans l?affichage du code HTML dans les diff?rents navigateurs. Cette rubrique explique comment d?boguer les compl?ments en ex?cution sur iPad ou sur Mac.</span><span class="sxs-lookup"><span data-stu-id="82bad-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac.</span></span> 

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="82bad-106">D?bogage avec l'inspecteur Web de Safari sur un Mac</span><span class="sxs-lookup"><span data-stu-id="82bad-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="82bad-107">Vous pouvez d?boguer un compl?ment Office ? l'aide de l'inspecteur Web de Safari.</span><span class="sxs-lookup"><span data-stu-id="82bad-107">You can debug an Office add-in using Safari Web Inspector.</span></span> 

<span data-ttu-id="82bad-108">Pour pouvoir d?boguer les compl?ments Office sur Mac, vous devez disposer de Mac OS High Sierra ET de Mac Office Version : 16.9.1 (Build 18012504) ou version ult?rieure.</span><span class="sxs-lookup"><span data-stu-id="82bad-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="82bad-109">Si vous n'avez pas de build Office pour Mac, vous pouvez en obtenir un en rejoignant notre [programme pour les d?veloppeurs Office 365](https://aka.ms/o365devprogram).</span><span class="sxs-lookup"><span data-stu-id="82bad-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="82bad-110">Pour commencer, ouvrez un terminal et r?glez la propri?t? `OfficeWebAddinDeveloperExtras` pour l'application Office concern?e en proc?dant comme suit?:</span><span class="sxs-lookup"><span data-stu-id="82bad-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="82bad-111">Ensuite, ouvrez l'application Office et ins?rez votre compl?ment.</span><span class="sxs-lookup"><span data-stu-id="82bad-111">Then, open the Office application and insert your add-in.</span></span> <span data-ttu-id="82bad-112">Cliquez avec le bouton droit sur le compl?ment et vous devriez voir une option **Inspecter l'?l?ment** dans le menu contextuel.</span><span class="sxs-lookup"><span data-stu-id="82bad-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span>  <span data-ttu-id="82bad-113">S?lectionnez cette option et l'inspecteur appara?tra, o? vous pouvez d?finir des points d'arr?t et d?boguer votre compl?ment.</span><span class="sxs-lookup"><span data-stu-id="82bad-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="82bad-114">Veuillez noter qu'il s'agit d'une fonctionnalit? exp?rimentale et il n'y a aucune garantie que nous allons la conserver dans les futures versions des applications Office.</span><span class="sxs-lookup"><span data-stu-id="82bad-114">Please note that this is an experimental feature and there are no guarantees that we will preserve this functionality in future versions of Office applications.</span></span>

## <a name="debugging-with-vorlonjs-on-a-ipad-or-mac"></a><span data-ttu-id="82bad-115">D?bogage avec Vorlon.JS sur iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="82bad-115">Debugging with Vorlon.JS on a iPad or Mac</span></span>

<span data-ttu-id="82bad-116">Pour d?boguer un compl?ment sur iPad ou Mac, vous pouvez utiliser Vorlon.JS, un d?bogueur pour pages Web similaire aux outils F12.</span><span class="sxs-lookup"><span data-stu-id="82bad-116">To debug an add-in on iPad or Mac, you can use Vorlon.JS, a debugger for web pages that is similar to the F12 tools.</span></span> <span data-ttu-id="82bad-117">Il est con?u pour fonctionner ? distance et d?boguer des pages web sur diff?rents appareils.</span><span class="sxs-lookup"><span data-stu-id="82bad-117">It is designed to work remotely and it enables you to debug web pages across different devices.</span></span> <span data-ttu-id="82bad-118">Pour plus d?informations, consultez le [site web de Vorlon](http://www.vorlonjs.com).</span><span class="sxs-lookup"><span data-stu-id="82bad-118">For more information, see the [Vorlon website](http://www.vorlonjs.com).</span></span>  


### <a name="install-and-set-up-vorlonjs"></a><span data-ttu-id="82bad-119">Installation et configuration de Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="82bad-119">Install and set up up Vorlon.JS on a Mac or iPad</span></span>  

1.  <span data-ttu-id="82bad-120">Connectez-vous au support en tant qu?administrateur.</span><span class="sxs-lookup"><span data-stu-id="82bad-120">Log on to the device as an administrator.</span></span>

2.  <span data-ttu-id="82bad-121">Installez [Node.js](https://nodejs.org) s?il n?est pas d?j? install?.</span><span class="sxs-lookup"><span data-stu-id="82bad-121">Install [Node.js](https://nodejs.org) if it isn't already installed.</span></span> 

3.  <span data-ttu-id="82bad-p105">Ouvrez une fen?tre **Terminal** et entrez la commande `npm i -g vorlon`. L?outil est install? dans le dossier `/usr/local/lib/node_modules/vorlon`.</span><span class="sxs-lookup"><span data-stu-id="82bad-p105">Open a **Terminal** window and enter the command `npm i -g vorlon`. The tool is installed to `/usr/local/lib/node_modules/vorlon`.</span></span>


### <a name="configure-vorlonjs-to-use-https"></a><span data-ttu-id="82bad-124">Configuration de Vorlon.JS pour une utilisation avec le protocole HTTPS</span><span class="sxs-lookup"><span data-stu-id="82bad-124">Configure Vorlon.JS to use HTTPS</span></span>

<span data-ttu-id="82bad-p106">Pour d?boguer une application ? l?aide de Vorlon.JS, ajoutez la balise `<script>` ? la page d?ouverture de l?application qui charge un script Vorlon.JS ? partir d?un emplacement connu (pour plus de d?tails, reportez-vous ? la proc?dure suivante). Si un compl?ment est s?curis? par SSL (HTTPS), tout script qu?il utilis doit ?tre h?berg? sur un serveur HTTPS, y compris le script Vorlon.JS. Par cons?quent, afin d?utiliser Vorlon.JS avec des compl?ments, vous devez le configurer pour qu?il se serve du protocole SSL.</span><span class="sxs-lookup"><span data-stu-id="82bad-p106">To debug an application using Vorlon.JS, you add a `<script>` tag to the opening page of the application that loads a Vorlon.JS script from a well-known location (for details, see the following procedure). If an add-in is SSL-secured (HTTPS), any scripts that it uses must be hosted from an HTTPS server, including the Vorlon.JS script. Therefore, you must configure Vorlon.JS to use SSL in order to use Vorlon.JS with add-ins.</span></span> 

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  <span data-ttu-id="82bad-128">Dans **Finder**, acc?dez ? `/usr/local/lib/node_modules/vorlon`, ouvrez le menu contextuel (en cliquant avec le bouton droit) du dossier `/Server`, puis s?lectionnez **Lire les informations**.</span><span class="sxs-lookup"><span data-stu-id="82bad-128">In **Finder**, go to `/usr/local/lib/node_modules/vorlon`, open the context menu for (right-click) the `/Server` folder, and then select **Get Info**.</span></span>

2.  <span data-ttu-id="82bad-129">Cliquez sur l?ic?ne en forme de cadenas dans le coin inf?rieur droit de la fen?tre **Informations sur le serveur** pour d?verrouiller le dossier.</span><span class="sxs-lookup"><span data-stu-id="82bad-129">Choose the padlock icon in the lower right corner of the **Server info** window to unlock the folder.</span></span>

3. <span data-ttu-id="82bad-130">Dans la section **Partage et permissions** de la fen?tre, d?finissez le **privil?ge** pour le groupe **personnel** sur **Lecture et ?criture**.</span><span class="sxs-lookup"><span data-stu-id="82bad-130">In the **Sharing and Permissions** section of the window, set the **Privilege** for the **staff** group to **Read & Write**.</span></span>

4. <span data-ttu-id="82bad-131">Cliquez ? nouveau sur l?ic?ne en forme de cadenas pour ***verrouiller ? nouveau*** le dossier.</span><span class="sxs-lookup"><span data-stu-id="82bad-131">Choose the padlock icon again to ***relock*** the folder.</span></span>

5. <span data-ttu-id="82bad-132">Dans **Finder**, d?veloppez le sous-dossier `/Server`, cliquez avec le bouton droit sur le fichier `config.json`, puis s?lectionnez **Lire les informations**.</span><span class="sxs-lookup"><span data-stu-id="82bad-132">Back in **Finder**, expand the `/Server` subfolder, right-click the file `config.json`, and then select **Get Info**.</span></span>

6. <span data-ttu-id="82bad-p107">Dans la fen?tre **config.json info**, modifiez les privil?ges du fichier de la m?me fa?on que pour le dossier `/Server` parent. Verrouillez ? nouveau le dossier et fermez la fen?tre.</span><span class="sxs-lookup"><span data-stu-id="82bad-p107">In the **config.json info** window, change the privileges of the file exactly the way you did for its parent `/Server` folder. Be sure to relock and close the window.</span></span>

7. <span data-ttu-id="82bad-p108">Dans **Finder**, cliquez avec le bouton droit sur le fichier `config.json`, s?lectionnez **Ouvrir avec**, puis **TextEdit**. Le fichier s?ouvre dans un ?diteur de texte.</span><span class="sxs-lookup"><span data-stu-id="82bad-p108">Back in **Finder**, right-click the file `config.json`, select **Open with**, and then select **TextEdit**. The file opens in a text editor.</span></span>

8. <span data-ttu-id="82bad-137">D?finissez la valeur de la propri?t? **useSSL** sur `true`.</span><span class="sxs-lookup"><span data-stu-id="82bad-137">Change the value of the **useSSL** property to `true`.</span></span>

9. <span data-ttu-id="82bad-p109">Dans la section **Modules**, recherchez le module ayant pour **ID** `OFFICE` et pour **nom** `Office Addin`. Si la valeur de la propri?t? **enabled** pour le module n?est pas d?j? d?finie sur `true`, d?finissez-la sur `true`.</span><span class="sxs-lookup"><span data-stu-id="82bad-p109">In the **plugins** section, find the plugin with the **id** of `OFFICE` and the **name** of `Office Addin`. If the **enabled** property for the plug-in is not already `true`, set it to `true`.</span></span>

10. <span data-ttu-id="82bad-140">Enregistrez le fichier et fermez l??diteur.</span><span class="sxs-lookup"><span data-stu-id="82bad-140">Save the file and close the editor.</span></span>

11. <span data-ttu-id="82bad-141">Dans **Finder**, acc?dez ? `/usr/local/lib/node_modules/vorlon`, cliquez avec le bouton droit sur le sous-dossier `Server`, et s?lectionnez **Nouveau terminal au dossier**.</span><span class="sxs-lookup"><span data-stu-id="82bad-141">In **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
12. <span data-ttu-id="82bad-p110">Dans la fen?tre **Terminal**, entrez `sudo vorlon`. Vous ?tes invit? ? saisir le mot de passe de l?administrateur. Le serveur Vorlon d?marre. Laissez la fen?tre **Terminal** ouverte.</span><span class="sxs-lookup"><span data-stu-id="82bad-p110">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

13. <span data-ttu-id="82bad-p111">Ouvrez une fen?tre de navigateur et acc?dez ? `https://localhost:1337`, qui est l?interface de Vorlon.JS. Lorsque vous y ?tes invit?, s?lectionnez **Toujours** pour approuver le certificat de s?curit?.</span><span class="sxs-lookup"><span data-stu-id="82bad-p111">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface. When prompted, choose **Always** to trust the security certificate.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="82bad-p112">Si aucune fen?tre d?invite n?appara?t, il se peut que vous deviez approuver le certificat manuellement. Le fichier de certificat est le suivant : `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Suivez la proc?dure ci-dessous. Si vous rencontrez des probl?mes, consultez l?aide de Macintosh ou iPad.</span><span class="sxs-lookup"><span data-stu-id="82bad-p112">If you are not prompted, you might need to trust the certificate manually. The certificate file is `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Try the following steps. If you have trouble, consult Macintosh or iPad help.</span></span> 
    >
    > 1. <span data-ttu-id="82bad-152">Fermez la fen?tre du navigateur et, dans la fen?tre **Terminal** en cours d?ex?cution sur le serveur Vorlon, utilisez le raccourci Ctrl+C pour arr?ter le serveur.</span><span class="sxs-lookup"><span data-stu-id="82bad-152">Close the browser window and in the **Terminal** window that is running the Vorlon server, use Control-C to stop the server.</span></span>
    > 2. <span data-ttu-id="82bad-p113">Dans **Finder**, cliquez avec le bouton droit sur le fichier `server.crt` et s?lectionnez **Trousseaux d?acc?s**. La fen?tre **Trousseaux d?acc?s** s?ouvre.</span><span class="sxs-lookup"><span data-stu-id="82bad-p113">In **Finder**, right-click the `server.crt` file and select **Keychain Access**. The **Keychain Access** window opens.</span></span>
    > 3. <span data-ttu-id="82bad-p114">Dans la liste **Trousseaux** sur la gauche, s?lectionnez **Connexion** si l?option n?est pas d?j? s?lectionn?e, puis choisissez **Certificats** dans la section **Cat?gorie**. Le certificat **localhost** figure dans la liste.</span><span class="sxs-lookup"><span data-stu-id="82bad-p114">In the **Keychains** list on the left, select **login** if it is not already selected, and then select **Certificates** in the **Category** section. The certificate **localhost** is listed.</span></span>
    > 4. <span data-ttu-id="82bad-p115">Cliquez avec le bouton droit sur le certificat **localhost** et s?lectionnez **Lire les informations**. Une fen?tre **localhost** s?ouvre.</span><span class="sxs-lookup"><span data-stu-id="82bad-p115">Right-click the certificate **localhost** and select **Get Info**. A **localhost** window opens.</span></span>
    > 5. <span data-ttu-id="82bad-159">Dans la section **Approuver**, ouvrez le s?lecteur nomm? **Lors de l?utilisation de ce certificat** et s?lectionnez **Toujours approuver**.</span><span class="sxs-lookup"><span data-stu-id="82bad-159">In the **Trust** section, open the selector labeled **When using this certificate** and select **Always Trust**.</span></span> 
    > 6. <span data-ttu-id="82bad-p116">Fermez la fen?tre **localhost**. Si l?action r?ussit, une croix blanche dans un cercle bleu appara?t sur l?ic?ne du certificat **localhost** dans la fen?tre **Trousseaux d?acc?s**.</span><span class="sxs-lookup"><span data-stu-id="82bad-p116">Close the **localhost** window. If the action was successful, the **localhost** certificate in the **Keychain Access** window has a white cross in a blue circle on its icon.</span></span>


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a><span data-ttu-id="82bad-162">Configuration du compl?ment pour le d?bogage Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="82bad-162">Configure the add-in for Vorlon.JS debugging</span></span>

1. <span data-ttu-id="82bad-163">Ajoutez la balise de script suivante ? la section `<head>` du fichier home.html (ou fichier HTML principal) de votre compl?ment :</span><span class="sxs-lookup"><span data-stu-id="82bad-163">Add the following script tag to the `<head>` section of the home.html file (or main HTML file) of your add-in:</span></span>

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>    
    ```  

2. <span data-ttu-id="82bad-164">D?ployez l?application web du compl?ment sur un serveur web accessible ? partir de l?ordinateur Mac ou de l?iPad, tel qu?un site web Azure.</span><span class="sxs-lookup"><span data-stu-id="82bad-164">Deploy the add-in web application to a web server that is accessible from the Mac or iPad, such as an Azure website.</span></span> 

3. <span data-ttu-id="82bad-165">Mettez ? jour l?URL du compl?ment ? tous les emplacements o? elle appara?t dans le manifeste du compl?ment.</span><span class="sxs-lookup"><span data-stu-id="82bad-165">Update the URL of the add-in in all the places where the URL appears in the add-in manifest.</span></span>

4. <span data-ttu-id="82bad-166">Copiez le manifeste du compl?ment dans le dossier suivant sur l?ordinateur Mac ou l?iPad : `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, o? *{host_name}* est Word, Excel, PowerPoint ou Outlook.</span><span class="sxs-lookup"><span data-stu-id="82bad-166">Copy the add-in manifest to the following folder on the Mac or iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, where *{host_name}* is Word, Excel, PowerPoint, or Outlook.</span></span>


### <a name="inspect-an-add-in-in-vorlonjs"></a><span data-ttu-id="82bad-167">V?rification d?un compl?ment dans Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="82bad-167">Inspect an add-in in Vorlon.JS</span></span>

1. <span data-ttu-id="82bad-168">Si le serveur Vorlon n?est pas en cours d?ex?cution, dans **Finder**, acc?dez ? `/usr/local/lib/node_modules/vorlon`, puis cliquez avec le bouton droit sur le sous-dossier `Server` et s?lectionnez **Nouveau terminal au dossier**.</span><span class="sxs-lookup"><span data-stu-id="82bad-168">If the Vorlon server is not running, in **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
2.  <span data-ttu-id="82bad-p117">Dans la fen?tre **Terminal**, entrez `sudo vorlon`. Vous ?tes invit? ? saisir le mot de passe de l?administrateur. Le serveur Vorlon d?marre. Laissez la fen?tre **Terminal** ouverte.</span><span class="sxs-lookup"><span data-stu-id="82bad-p117">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

3.  <span data-ttu-id="82bad-173">Ouvrez une fen?tre de navigateur et acc?dez ? `https://localhost:1337`, qui est l?interface de Vorlon.JS.</span><span class="sxs-lookup"><span data-stu-id="82bad-173">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface.</span></span>

4. <span data-ttu-id="82bad-p118">Chargez une version test du compl?ment. S?il s?agit d?un compl?ment pour Excel, PowerPoint ou Word, chargez une version test en suivant les ?tapes d?crites dans la rubrique relative au [chargement d?une version test d?un compl?ment Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md). S?il s?agit d?un compl?ment Outlook, chargez une version de test en suivant les ?tapes d?crites dans la rubrique relative au [chargement d?une version test de compl?ments Outlook ? des fins de test](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing). Si le compl?ment n?utilise pas les commandes du compl?ment, il s?ouvre automatiquement. Sinon, cliquez sur le bouton d?ouverture du compl?ment. En fonction de la version de l?application h?te d?Office, vous trouverez le bouton sur l?onglet **Accueil** ou sur l?onglet **Compl?ment**.</span><span class="sxs-lookup"><span data-stu-id="82bad-p118">Sideload the add-in. If it is for Excel, PowerPoint, or Word, sideload it as described in [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md). If it is an Outlook add-in, sideload it as described in [Sideload Outlook Add-ins for testing](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing). If the add-in does not use add-in commands, it will open immediately. Otherwise, choose the button to open the add-in. Depending on the build of the Office host application, the button will be on either the **Home** tab or an **Add-in** tab.</span></span>

<span data-ttu-id="82bad-180">Le compl?ment appara?t dans la liste des clients dans Vorlon.JS (sur la gauche dans l?interface de Vorlon.JS) en tant que **{Syst?me d?exploitation} - n**, pour un nombre *n*, et o? *{Syst?me d?exploitation}* correspond au type d?appareil (par exemple, ? Macintosh ?).</span><span class="sxs-lookup"><span data-stu-id="82bad-180">The add-in will show up in the list of Clients in Vorlon.JS (on the left side of the Vorlon.JS interface) as **{OS} - n**, for some number *n*, and where *{OS}* is the device type, such as "Macintosh".</span></span> 

![Capture d??cran de l?interface Vorlon.js](../images/vorlon-interface.png)

<span data-ttu-id="82bad-p119">L?outil Vorlon dispose d?une vari?t? de plug-ins. Les plug-ins actuellement activ?s apparaissent sous forme d?onglets dans la partie sup?rieure de l?interface de l?outil. (Vous pouvez en activer davantage en cliquant sur l?ic?ne en forme d?engrenage sur la gauche.) Ces plug-ins sont semblables aux fonctions disponibles dans les outils F12. Par exemple, vous pouvez mettre en surbrillance les ?l?ments DOM, ex?cuter des commandes, etc. Pour plus d?informations, reportez-vous ? la page relative ? la [documentation principale sur les plug-ins Vorlon](http://vorlonjs.com/documentation/#console).</span><span class="sxs-lookup"><span data-stu-id="82bad-p119">The Vorlon tool has a variety of plug-ins. The ones that are currently enabled appear as tabs at the top of the tool. (You can enable more plug-ins by choosing the gears icon on the left.) These plug-ins are  similar to the functions in F12 tools. For example, you can highlight DOM elements, execute commands, and more. For more details, see [Vorlon Documentation Core Plugins](http://vorlonjs.com/documentation/#console)</span></span> 

<span data-ttu-id="82bad-p120">Un plug-in **Compl?ment Office** permet d?ajouter des fonctionnalit?s suppl?mentaires pour Office.js, telles que l?exploration du mod?le objet, l?ex?cution d?appels Office.js et la lecture des valeurs des propri?t?s de l?objet. Pour plus d?informations, reportez-vous ? l?article relatif ? l?utilisation du [plug-in VorlonJS pour d?boguer un compl?ment Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span><span class="sxs-lookup"><span data-stu-id="82bad-p120">An **Office Addin** plug-in adds extra capabilities for Office.js, such as exploring the object model, executing Office.js calls, and reading the values of object properties. For instructions, see [VorlonJS plugin for debugging Office Add-in](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span></span>

> [!NOTE]
> <span data-ttu-id="82bad-188">il n?existe aucun moyen de d?finir des points d?arr?t dans Vorlon.JS.</span><span class="sxs-lookup"><span data-stu-id="82bad-188">There is no way to set break points in Vorlon.JS.</span></span>


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a><span data-ttu-id="82bad-189">Effacement du cache de l?application Office sur un ordinateur Mac ou un iPad</span><span class="sxs-lookup"><span data-stu-id="82bad-189">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="82bad-p121">Les compl?ments sont souvent mis en cache dans Office pour Mac, pour des raisons de performances. En r?gle g?n?rale, vous pouvez effacer le cache en rechargeant le compl?ment. En pr?sence de plusieurs compl?ments dans le m?me document, il se peut que le processus d?effacement automatique du cache lors du rechargement ne fonctionne pas syst?matiquement.</span><span class="sxs-lookup"><span data-stu-id="82bad-p121">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span> 

<span data-ttu-id="82bad-193">Sur un ordinateur Mac, vous pouvez effacer le cache manuellement en supprimant tous les ?l?ments contenus dans le dossier `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="82bad-193">On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

<span data-ttu-id="82bad-p122">Sur un iPad, vous pouvez appeler `window.location.reload(true)` ? partir de JavaScript dans le compl?ment pour forcer le rechargement. Vous pouvez ?galement choisir de r?installer Office.</span><span class="sxs-lookup"><span data-stu-id="82bad-p122">On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>
