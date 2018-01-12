
# <a name="walkthrough-creating-your-first-lab-for-office-mix"></a>Procédure : Création de votre premier atelier pour Office Mix
Créez votre premier atelier LabsJS en suivant une procédure pas à pas.



Dans cette procédure pas à pas, vous allez créer un atelier LabsJS simple à partir de zéro. Votre atelier sera un simple questionnaire vrai/faux fournissant une seule question. 

Plutôt que de commencer avec un modèle de projet Visual Studio, vous allez commencer avec trois fichiers vides (cela montre à quel point la création d’un atelier est directe) : 


- TrueFalse.html (html5)
    
- TrueFalse.js
    
- TrueFalse.css
    
Vous pouvez utiliser l’éditeur de code de votre choix pour modifier ces fichiers, car nous ne nous basons pas sur un modèle Visual Studio. En fait, le fichier HTML n’est pas très important, et si vous le souhaitez vous pouvez simplement copier/coller le balisage HTML indiqué dans les didacticiels. Toutefois, il doit s’agir de HTML5, assurez-vous donc que la déclaration DOCTYPE est  `<!DOCTYPE html>`. Le fichier CSS est facultatif. Tous les éléments importants sont réalisés dans le fichier JavaScript (.js), TrueFalse.js. La procédure pas à pas va aborder quatre principales caractéristiques d’atelier :

- Configuration (connexion à l’hôte)
    
- Changements de mode (entre le mode d’édition et le mode d’affichage)
    
- Modification de l’atelier
    
- Exécution de l’atelier
    

 **Remarque**  
 ---
 Le fichier labhost.html est exécuté sur un serveur web et fournit l’environnement d’hébergement pour le développement et les tests de l’atelier. Cela simplifie grandement le développement de l’atelier. Voir [Prise en main de LabsJS pour Office Mix](get-started-with-labsjs-for-office-mix.md) pour plus d’informations sur la configuration de votre environnement de développement.<br/><br/>

Enfin, vous pouvez voir les fichiers JavaScript achevés (TrueFalse.js) parmi les fichiers distribués avec ce SDK. Vous trouverez ci-dessous une procédure pas à pas du processus de codage.

## <a name="connecting-to-the-lab-host"></a>Connexion à l’hôte de l’atelier

Dans cet environnement, les ateliers sont en mesure de fonctionner soit avec l’hôte d’atelier (pour le développement et les tests), soit avec l’hôte d’exécution par défaut fourni par l’hôte Office.js. La fonction d’ouverture utilise ensuite une expression simple if/else pour tester lequel de ces contextes d’hébergement s’applique.


```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```

L’objet  **PostMessageLabHost** s’exécute dans l’environnement de développement labhost.html, tandis que dans l’environnement de production, l’atelier s’exécute dans PowerPoint/Office Mix à l’aide de l’élément **OfficeJSLabHost**.

Ensuite, créez une méthode d’assistance pour créer un rappel dont la tâche consiste à résoudre ou à rejeter un objet différé jQuery que vous transmettez. Utilisez la méthode  **createCallback** pour passer des promesses jQuery aux rappels définis par labs.js.




```js
function createCallback(deferred) {
    return function (err, data) {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```

Nous créons également une méthode d’assistance pour récupérer la configuration de l’atelier pour une question et une réponse données.




```js
function getConfiguration(question, answer) {
    var choiceComponent = {
        name: question,
        type: Labs.Components.ChoiceComponentType,
        timeLimit: 0,
        maxAttempts: 1,
        choices: [
            { id: "0", name: "True", value: "True" },
            { id: "1", name: "False", value: "False" }],
        maxScore: 1,
        hasAnswer: true,
        answer: answer ? "0" : "1",
        values: null,
        secure: false,
        data: null
    };

    return {
        appVersion: { major: 0, minor: 1 },
        components: [choiceComponent],
        name: question,
        timeline: null,
        analytics: null
    };
}
```


## <a name="mode-changes"></a>Changement de mode

Un atelier est toujours dans l’un des deux états ou modes :  **view** et **edit**. Par conséquent, nous devons capturer et maintenir l’état et le comportement pour le questionnaire. Nous allons créer une classe à cet effet.


```js
var TrueFalseQuiz = (function () {
    /**
     * Constructor - takes in the starting mode.
     */
    function TrueFalseQuiz(mode) {
        var self = this;        
        self._modeSwitchP = $.when();
        self._labInstance = null;
        self._labEditor = null;        
      /**
       * Listen for mode changed events and 
       * then switch accordingly. Also set the initial mode state.
       */
        Labs.on(Labs.Core.EventTypes.ModeChanged, function (modeChangedEvent) {
            self.switchUserMode(Labs.Core.LabMode[modeChangedEvent.mode]);
        });
        this.switchUserMode(mode);        
    }
```

En outre, nous fournissons une méthode d’assistance dont le travail est de mettre à jour l’interface utilisateur du questionnaire selon que la question (en d’autres termes, la « soumission ») est correcte ou incorrecte.




```js
    TrueFalseQuiz.prototype._showResults = function(correct) {
        $("#submit-button").removeClass("btn-default");
        $("#submit-button").addClass(correct ? "btn-success" : "btn-danger");
        $("#submit-button").text(correct ? "Correct!" : "Incorrect");

        $("#submit-button").prop("disabled", true);
        $("input:radio[name='quizAnswers']").prop("disabled", true);
    };
```

Nous avons également besoin d’une fonction pour passer d’un mode à l’autre.




```js
TrueFalseQuiz.prototype.switchUserMode = function (mode) {
        var self = this;

        // Wait for any previous mode switch to complete before performing the new one
        self._modeSwitchP = self._modeSwitchP.then(function () {
            var switchedStateDeferred = $.Deferred();

            // Clean up any variables associated with the previous mode.
            if (self._labInstance) {
                $("#quiz-view-form").off("submit");
                self._labInstance.done(createCallback(switchedStateDeferred));
            } else if (self._labEditor) {
                self._unbindFromEditUpdates();
                self._labEditor.done(createCallback(switchedStateDeferred));
            } else {
                switchedStateDeferred.resolve();
            }

            // After the cleanup occurs, switch to the new mode.
            return switchedStateDeferred.promise().then(function () {
                self._labEditor = null;
                self._labInstance = null;

                if (mode === Labs.Core.LabMode.Edit) {
                    return self._switchToEditMode();
                } else {
                    return self._switchToViewMode();
                }
            });
        });

        // Display an error if it occurs.
        self._modeSwitchP.fail(function (error) {
            // ... error handling ...
        });
    };
```

La fonction suivante met à jour la configuration du questionnaire en fonction des événements de modification que nous avons reçus de la part de l’interface utilisateur.




```js
    TrueFalseQuiz.prototype._updateConfigurationFromUI = function () {
        var question = $("#question-edit").val();
        var answerIsTrue = $("input:radio[name='answerValue']:checked").val() === "true";

        this._updateConfiguration(question, answerIsTrue, true, function (err) {
            if (err) {
                // show error
            }
        });
    };
```

Ensuite, nous mettons à jour les données de configuration de l’atelier stockées sur le serveur en fonction des questions et réponses données.




```js
    TrueFalseQuiz.prototype._updateConfiguration = function (question, answer, serialize, callback) {
        var configuration = getConfiguration(question, answer);

        if (serialize) {
            this._labEditor.setConfiguration(configuration, callback);
        } else {
            callback(null, null);
        }
    };
```

Ensuite, nous avons une fonction qui lie les mises à jour effectuées dans l’atelier en mode d’édition aux modifications de configuration que nous avons apportées, puis nous utilisons un code de séparation provenant des gestionnaires de modification précédemment liés.




```js
    TrueFalseQuiz.prototype._bindToEditUpdates = function () {
        var self = this;

        // Listen for the question changing
        $("#question-edit").on("input propertychange paste", function () {
            self._updateConfigurationFromUI();
        });

        $('input[name="answerValue"]').on("change", function (e) {
            self._updateConfigurationFromUI();
        });
    };
```




```js
    TrueFalseQuiz.prototype._unbindFromEditUpdates = function () {
        $("#question-edit").off("input propertychange paste");
        $('input[name="answerValue"]').off("change");
    };
```

Nous arrivons maintenant à un élément clé de la section, autrement dit aux méthodes permettant de passer d’un mode à l’autre entre les modes d’affichage et d’édition. Commençons par le passage du mode d’affichage au mode d’édition.




```js
    TrueFalseQuiz.prototype._switchToEditMode = function () {
        var self = this;
        var editLabDeferred = $.Deferred();

        // Make the Labs.js API call to edit the lab.
        Labs.editLab(createCallback(editLabDeferred));

        return editLabDeferred.promise().then(function (labEditor) {            
            self._labEditor = labEditor;

            // Retrieve any existing configuration from the lab editor.
            var configurationDeferred = $.Deferred();
            labEditor.getConfiguration(createCallback(configurationDeferred));

            return configurationDeferred.promise().then(function (configuration) {
                var configurationReadyDeferred = $.Deferred();

                // Get the question and answer values if they exist. 
                //Otherwise use the defaults.
                var question = configuration !== null ? configuration.components[0].name : "";
                var answerIsTrue = configuration !== null ? configuration.components[0].answer === "0" : true;

                // Update the lab configuration based on the question and answer.
                self._updateConfiguration(
                    question,
                    answerIsTrue,
                    configuration === null,
                    createCallback(configurationReadyDeferred));

                // Update the UI based on the question and answer.
                $("#question-edit").val(question);
                $('input[name="answerValue"][value="' + answerIsTrue + '"]').prop('checked', true);

                // Bind to changes.
                self._bindToEditUpdates();

                // Flip over the UI.
                $("#quiz-editor").removeClass("hidden");
                $("#quiz-view").addClass("hidden");

                return configurationReadyDeferred.promise();
            });
        });
    };
```

Maintenant, nous allons parler du passage du mode d’édition au mode d’affichage.




```js
    TrueFalseQuiz.prototype._switchToViewMode = function () {
        var self = this;
        var takeLabDeferred = $.Deferred();

        // Call the labs.js API to start taking the lab.
        Labs.takeLab(createCallback(takeLabDeferred));

        return takeLabDeferred.promise().then(function (labInstance) {
            self._labInstance = labInstance;

            // Get the choice component instance that will be generated
            // from the choice component we saved when editing the lab.
            var choiceComponentInstance = self._labInstance.components[0];

            // Get the attempts associated with that choice component.
            var attemptsDeferred = $.Deferred();
            choiceComponentInstance.getAttempts(createCallback(attemptsDeferred));
            var attemptP = attemptsDeferred.promise().then(function (attempts) {
                // See if we already had started an attempt against 
                // the problem. If not create one.
                var currentAttemptDeferred = $.Deferred();
                if (attempts.length > 0) {
                    currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
                } else {
                    choiceComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
                }

                return currentAttemptDeferred.then(function (currentAttempt) {
                    var resumeDeferred = $.Deferred();

                    // After we have the attempt, mark that we are resuming
                    // it as well. This will note the resumption time
                    // in the lab activity log.
                    currentAttempt.resume(createCallback(resumeDeferred));
                    return resumeDeferred.promise().then(function () {
                        return currentAttempt;
                    });
                });
            });

            return attemptP.promise().then(function (attempt) {
                // Store off the latest attempt for later use.
                self._currentAttempt = attempt;

                // Update the question field of the view UI.
                $("#question-view").text(choiceComponentInstance.component.name);

                // Determine whether the quiz has already been taken
                // and update the UI accordingly.
                var submissions = attempt.getSubmissions();
                if (submissions.length > 0) {
                    var correctAttempt = submissions[submissions.length - 1].result.score === 1;
                    var submissionValue = submissions[submissions.length - 1].answer.answer === "0";
                    $('input[name="quizAnswers"][value="' + submissionValue + '"]').prop('checked', true);
                    self._showResults(correctAttempt);
                } else {
                    $("#submit-button").removeClass("btn-success btn-danger"    );
                    $("#submit-button").addClass("btn-default");
                    $("#submit-button").text("Submit");
                    $("#submit-button").prop("disabled", false);
                    $("input:radio[name='quizAnswers']").prop("disabled", false);
                }                

                // Hook up the form submit button and then
                // grade the attempt when it is selected.
                $("#quiz-view-form").on("submit", function (e) {
                    e.preventDefault();
                    
                    // Get the checked value and see whether the choice
                    // was true or false - map back to our choice fields.
                    var submission = $("input:radio[name='quizAnswers']:checked").val() === "true" ? "0" : "1";

                    // Grade against the stored answer.
                    var correct = choiceComponentInstance.component.answer === submission;

                    // Submit the attempt with the labs.js API.
                    attempt.submit(
                        new Labs.Components.ChoiceComponentAnswer(submission),
                        new Labs.Components.ChoiceComponentResult(correct ? 1 : 0, true),
                        function (err) {
                            if (err) {
                                // Error
                            }
                        });

                    // And finally update the UI.
                    self._showResults(correct);
                });

                // And make the view UI visible.
                $("#quiz-editor").addClass("hidden");
                $("#quiz-view").removeClass("hidden");
            });
        });
    };

    return TrueFalseQuiz;
})();
```

Enfin, une fois que vous êtes connecté à l’hôte et que le document est prêt, vous pouvez démarrer le questionnaire.




```js
$(document).ready(function () {
    Labs.connect(function (err, connectionResponse) {
        if (err) {
            // ... error handling goes here ...
            return;
        }

        // Start up the true/false quiz.
        var trueFalseQuiz = new TrueFalseQuiz(connectionResponse.mode);
    });
});
```


## <a name="additional-resources"></a>Ressources supplémentaires
<a name="bk_addresources"> </a>


- [Compléments Office Mix](office-mix-add-ins.md)
    
