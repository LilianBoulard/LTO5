"""

Ce script a pour rôle de sauvegarder les mails de rapports trouvés dans une boite mail.

"""

import sys

if sys.version_info[0] != 3: # Si la version de Python n'est pas la 3.x
    input('Ce script doit être exécuté avec Python 3')
    exit(0)

import os, datetime

# On importe le fichier de traitement de bandes (IMPORTANT: Le nom doit correspondre au nom du script !)
import Traitement_bandes

"""
NOTE : pour info sur la méthode SaveAs :
OlSaveAsType = {
    "olTXT": 0,
    "olRTF": 1,
    "olTemplate": 2,
    "olMSG": 3,
    "olDoc": 4,
    "olHTML": 5,
    "olVCard": 6,
    "olVCal": 7,
    "olICal": 8
}
"""



# ================================
#  Import des modules nécessaires
# ================================



try:
    import win32com.client
except ImportError:
    if Traitement_bandes.install_module("pypiwin32"):
        import win32com.client
    else:
        exit()



# ======
#  Code
# ======



def main():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Sélectionne l'application Outlook et son API

    for store in outlook.Stores: # Pour chaque boite mail trouvée dans Outlook
        if Traitement_bandes.verbose == True: # Si le debug est activé
            # On écrit le nom de la boite mail en cours de traitement
            print(store)

        if str(store) == Traitement_bandes.nomBoite: # Si le nom de la boite mail est "frhelpdesk@iriworldwide.com"
            # On sélectionne le dossier 6, qui est la boite de réception
            inbox = store.GetDefaultFolder(6)
            # Et récupère les mails du dossier.
            messages = inbox.Items

            # On initialise ue drapeau sur Faux. Il permet de savoir si un mail valide a été trouvé
            drapeau = False

            for message in messages: # Pour chaque message dans la boite
                # Si l'objet du mail est "Bandes_LTO5_FREEPOOL_COFFRE" ou "Bandes_LTO5_FREEPOOL"
                if message.Subject in Traitement_bandes.objets:
                    # Récupère sa date d'envoi, et la convertit en objet datetime (pour pouvoir être manipulé plus facilement pas la suite)
                    temps = datetime.datetime.strptime(str(message.SentOn)[:-6], '%Y-%m-%d %H:%M:%S')
                    # Transforme le datetime en date "lisible" (string), au format "formatDate" ; cette variable se trouve dans l'en-tête
                    tempsDate = temps.strftime(Traitement_bandes.formatDate)

                    # Détermine le nom du dossier dans lequel seront sauvegardés les mails en fonction de leur date de réception (au format "MoisAnnée" -> "Janvier2000")
                    dossier = (Traitement_bandes.mois[temps.month - 1] + str(temps.year)) + '/'
                    repertoire = Traitement_bandes.dossierBandes + dossier

                    try:
                        # Essaye de créer le répertoire où seront stockées les bandes
                        os.mkdir(repertoire)
                    except FileExistsError: # Si le répertoire existe déjà
                        # On ignore l'erreur
                        pass

                    # On construit le nom complet du fichier
                    fichier = message.Subject + '-' + tempsDate + Traitement_bandes.extension
                    # Et on l'enregistre. 5 correspond au type olHTML (Voir la note plus haut pour plus d'informations).
                    message.SaveAs(repertoire + fichier, 5)

                    print('Mail {0} enregistré dans le dossier {1}'.format(fichier, repertoire))

                    # Si un mail est trouvé, on passe le drapeau en vrai
                    drapeau = True

            # Si aucun mail n'a été trouvé
            if drapeau == False:
                input('Aucun fichier de bandes trouvé. Appuyez sur une entrée pour fermer le programme...')

            # Si un mail a été trouvé
            else:
                ch = input('Exécuté avec succès. Lancer le script de traitement des bandes ? [Oui/Non]')
                # Si la réponse donnée se trouve dans la liste des choix positifs
                if ch.lower() in Traitement_bandes.choixPositifs:
                    # On appelle le main() du script pour lancer le traitement
                    Traitement_bandes.main(Traitement_bandes.dossierBandes)
                else:
                    input('Réponse négative. Appuyez sur entrée pour fermer le programme...')
                    exit(0)



if __name__ == '__main__':
    main()
