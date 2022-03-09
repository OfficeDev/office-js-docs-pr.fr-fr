> [!NOTE]
> Outlook sur Windows : si vous exécutez votre add-in à partir de localhost et que vous voyez l’erreur « Nous sommes désolés, nous n’avons pas pu accéder à *{your-add-in-name-here}*. Assurez-vous que vous avez une connexion réseau. Si le problème persiste, veuillez essayer à nouveau plus tard. », vous devrez peut-être activer une exemption de bouclisation.
>
> 1. Fermez Outlook.
> 1. Ouvrez **le Gestionnaire des tâches** et assurez-vous que **le processusmsoadfsb.exe'est** pas en cours d’exécution.
> 1. Définissez [l’exemption de bouclisation](/previous-versions/windows/apps/hh780593(v=win.10)?redirectedfrom=MSDN) dans une invite avec élévation de élévation de droits.
>     - Si vous utilisez et portez `https://localhost` 3000 (configuration par défaut), exécutez la commande suivante.
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_https___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>     - Si vous utilisez et portez `http://localhost` 3000, exécutez la commande suivante.
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>
>      **Remarque** : si vous n’utilisez pas le port par défaut 3000, remplacez-le dans la commande par votre numéro de port réel.
> 1. Redémarrez Outlook.
