
<span data-ttu-id="393b1-101">Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="393b1-101">Complete the following steps to start the local web server and sideload your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="393b1-102">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="393b1-102">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="393b1-103">Si vous êtes invité à installer un certificat après avoir exécuté une des commandes suivantes, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="393b1-103">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

> [!TIP]
> <span data-ttu-id="393b1-104">Si vous testez votre complément sur Mac, exécutez la commande suivante avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="393b1-104">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="393b1-105">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="393b1-105">When you run this command, the local web server will start.</span></span>
>
> ```command&nbsp;line
> npm run dev-server
> ```

- <span data-ttu-id="393b1-106">Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="393b1-106">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="393b1-107">Cela a pour effet de démarrer le serveur web local (s’il n’est pas déjà en cours d’exécution) et d’ouvrir Excel avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="393b1-107">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="393b1-108">Pour tester votre complément dans Excel sur le web, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="393b1-108">To test your add-in in Excel on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="393b1-109">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="393b1-109">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="393b1-110">Pour utiliser votre complément, ouvrez un nouveau document dans Excel sur le web, puis chargez la version test de votre complément en suivant les instructions de l’article relatif au [chargement de la version test des compléments Office dans Office sur le web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="393b1-110">To use your add-in, open a new document in Word Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

