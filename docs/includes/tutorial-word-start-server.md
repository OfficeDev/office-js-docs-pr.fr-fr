<span data-ttu-id="4298f-101">Si le serveur Web local est déjà en cours d’exécution et que votre complément est déjà chargé dans Word, passez à l’étape 2.</span><span class="sxs-lookup"><span data-stu-id="4298f-101">If the local web server is already running and your add-in is already loaded in Word, proceed to step 2.</span></span> <span data-ttu-id="4298f-102">Dans le cas contraire, démarrez le serveur Web local et chargement votre complément :</span><span class="sxs-lookup"><span data-stu-id="4298f-102">Otherwise, start the local web server and sideload your add-in:</span></span> 

- <span data-ttu-id="4298f-103">Pour tester votre complément dans Word, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="4298f-103">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="4298f-104">Cela démarre le serveur Web local (s’il n’est pas déjà en cours d’exécution) et ouvre Word avec votre complément chargé.</span><span class="sxs-lookup"><span data-stu-id="4298f-104">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="4298f-105">Pour tester votre complément dans Word sur le Web, exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="4298f-105">To test your add-in in Word on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="4298f-106">Lorsque vous exécutez cette commande, le serveur Web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="4298f-106">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="4298f-107">Pour utiliser votre complément, ouvrez un nouveau document dans Word sur le Web, puis chargement votre complément en suivant les instructions de [chargement des compléments Office dans Office sur le Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="4298f-107">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>
