CH2_STRINGS = {"en": { # dialog controls (label text + help text)
                      "label_lang": ("Language", ""),
                      "label_style": ("Style", ""),
                      "check_col_bg": ("Set background color", "Use the background color provided by the style."),
                      "check_linenb": ("Line numbering", "Active or deactivate line numbers."),
                      "pygments_ver": ("Build upon Pygments {}", ""),
                       # messages
                      "errlang": "Unsupported language.",
                      "errstyle": "Unknown style.",
                      "errsel1": "Unsupported selection.",
                      "errsel2": "Nothing to highlight.",
                      },

               "fr": { # contrôles de dilaogue (texte + texte d'aide)
                      "label_lang": ("Langage", ""),
                      "label_style": ("Style", ""),
                      "check_col_bg": ("Utiliser la couleur de fond", "Utiliser l'arrière-plan fourni par le style."),
                      "check_linenb": ("Numérotation des lignes", "Affiche ou masque les numéros de ligne."),
                      "pygments_ver": ("Basé sur Pygments {}", ""),
                       # messages
                      "errlang": "Langage non supporté.",
                      "errstyle": "Style inconnu.",
                      "errsel1": "Sélection invalide.",
                      "errsel2": "Rien à coloriser.",
                      }
              }


def getstrings(ctx):
    '''Get interface locale'''
    ps = ctx.ServiceManager.createInstanceWithContext("com.sun.star.util.PathSubstitution", ctx)
    vlang = ps.getSubstituteVariableValue("vlang")
    lang = vlang.split("-")[0]
    # print('vlang: {}'.format(vlang))
    # print('lang: {}'.format(lang))
    return CH2_STRINGS.get(lang, CH2_STRINGS["en"])

