
<span data-ttu-id="26e48-101">Pour démarrer le serveur web local et charger indépendamment votre complément, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="26e48-101">Complete the following steps to start the local web server and sideload your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="26e48-102">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="26e48-102">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="26e48-103">Si vous êtes invité à installer un certificat après avoir exécuté une des commandes suivantes, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="26e48-103">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

> [!TIP]
> <span data-ttu-id="26e48-104">Si vous testez votre complément sur Mac, exécutez la commande suivante avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="26e48-104">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="26e48-105">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="26e48-105">When you run this command, the local web server will start.</span></span>
>
> ```command&nbsp;line
> npm run dev-server
> ```

- <span data-ttu-id="26e48-106">Pour tester votre complément dans Excel, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="26e48-106">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="26e48-107">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas encore en cours d’exécution), et Excel s’ouvre avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="26e48-107">When you run this command, the local web server will start and Word will open with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="26e48-108">Pour tester votre complément dans Excel Online, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="26e48-108">To test your add-in in Excel Online, run the following command in the root directory of your project.</span></span> <span data-ttu-id="26e48-109">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="26e48-109">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="26e48-110">Pour utiliser votre complément, ouvrez un nouveau document dans Excel Online, puis chargez indépendamment votre complément en suivant les instructions fournies dans [Chargement de version test d’un complément Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).</span><span class="sxs-lookup"><span data-stu-id="26e48-110">To use your add-in, open a new document in Word Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).</span></span>

