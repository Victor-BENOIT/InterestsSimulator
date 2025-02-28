import pandas as pd
import matplotlib.pyplot as plt
import mplcursors



def save_to_excel(account, filename="simulation_investissement.xlsx"):
    """Enregistre les données du compte d'investissement dans un fichier Excel."""
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        # Sauvegarde des données principales
        account.get_data().to_excel(writer, sheet_name="Simulation", index=True)

        # Ajout des paramètres en feuille séparée
        params_df = pd.DataFrame({
            "Paramètre": ["Nom", "Montant initial", "Frais de gestion (annuels)",
                          "Croissance du marché (annuelle)", "Rendement des dividendes",
                          "Contribution mensuelle"],
            "Valeur": [account.name, account.start_amount, account.management_fee,
                       account.market_growth, account.dividend_yield, account.monthly_contribution]
        })
        params_df.to_excel(writer, sheet_name="Paramètres", index=False)

        print(f"✅ Données enregistrées dans {filename}")


def plot_investment(account, *columns):
    """
    Affiche l'évolution d'une ou plusieurs colonnes du compte d'investissement au cours du temps,
    avec un curseur interactif affichant les valeurs de toutes les courbes pour une abscisse donnée.

    :param account: Instance de InvestmentAccount
    :param columns: Colonnes à afficher en courbes (ex: "Somme totale", "Bénéfices", etc.)
    """
    data = account.get_data()

    # Création d'un index numérique temporaire pour éviter les erreurs
    numeric_index = list(range(len(data)))

    plt.figure(figsize=(12, 6))

    curves = {}  # Stocker les courbes pour l'interaction
    for col in columns:
        if col in data.columns:
            curves[col], = plt.plot(numeric_index, data[col], label=col)
        else:
            print(f"⚠️ La colonne '{col}' n'existe pas dans les données.")

    # Configuration des labels de l'axe X (une seule fois par an)
    plt.xticks(ticks=numeric_index[::12], labels=data.index[::12], rotation=45)

    plt.xlabel("Temps")
    plt.ylabel("Valeur (€)")
    plt.title(f"Évolution des investissements pour {account.name}")
    plt.legend()
    plt.grid(True)

    # Ajout du curseur interactif
    cursor = mplcursors.cursor(list(curves.values()), hover=True)

    @cursor.connect("add")
    def on_hover(sel):
        index = int(sel.index)  # Convertir l'index flottant en entier
        year = data.index[index].split()[-1]  # Extraire l'année depuis l'index

        values = [f"{col}: {data[col].iloc[index]:,.2f} €" for col in columns]
        text = f"Année: {year}\n" + "\n".join(values)

        sel.annotation.set_text(text)
        sel.annotation.get_bbox_patch().set_facecolor("white")
        sel.annotation.get_bbox_patch().set_alpha(0.8)

    plt.show()

def plot_multiple_accounts(*accounts, column="Somme totale"):
    """
    Affiche l'évolution d'une métrique spécifique pour plusieurs comptes d'investissement.

    :param accounts: Instances de InvestmentAccount
    :param column: Nom de la colonne à afficher (ex: "Somme totale", "Bénéfices", etc.)
    """
    plt.figure(figsize=(12, 6))
    curves = {}

    for account in accounts:
        data = account.get_data()
        numeric_index = list(range(len(data)))
        if column in data.columns:
            curves[account.name], = plt.plot(numeric_index, data[column], label=account.name)
        else:
            print(f"⚠️ La colonne '{column}' n'existe pas dans les données de {account.name}.")

    # Configuration des labels de l'axe X (une seule fois par an)
    if accounts:
        plt.xticks(ticks=numeric_index[::12], labels=data.index[::12], rotation=45)

    plt.xlabel("Temps")
    plt.ylabel("Valeur (€)")
    plt.title(f"Évolution de {column} pour plusieurs investissements")
    plt.legend()
    plt.grid(True)

    # Ajout du curseur interactif
    cursor = mplcursors.cursor(list(curves.values()), hover=True)

    @cursor.connect("add")
    def on_hover(sel):
        index = int(sel.index)
        year = data.index[index].split()[-1]
        values = [f"{account.name}: {account.get_data()[column].iloc[index]:,.2f} €" for account in accounts]
        text = f"Année: {year}\n" + "\n".join(values)
        sel.annotation.set_text(text)
        sel.annotation.get_bbox_patch().set_facecolor("white")
        sel.annotation.get_bbox_patch().set_alpha(0.8)

    plt.show()


class InvestmentAccount:
    def __init__(self, name, start_amount, management_fee, market_growth, dividend_yield, monthly_contribution):
        self.name = name
        self.start_amount = start_amount
        self.management_fee = management_fee
        self.market_growth = market_growth
        self.dividend_yield = dividend_yield
        self.monthly_contribution = monthly_contribution

        self.data = pd.DataFrame(columns=[
            "Somme totale",
            "Somme totale investie",
            "Bénéfices",
            "Frais de gestion",
            "Montants totaux de frais de gestion",
            "Évolution du marché",
            "Dividendes versés",
            "Contribution ajoutée"
        ])

        self.months = ["Janvier", "Février", "Mars", "Avril",
                       "Mai", "Juin", "Juillet", "Août",
                       "Septembre", "Octobre", "Novembre", "Décembre"]
        self.current_year = 2025

    def simulate(self, years):
        total_amount = self.start_amount
        total_invested = self.start_amount
        total_fees = 0
        total_fees_cumulative = 0  # Nouvelle variable pour accumuler les frais
        total_dividends = 0

        for year in range(1, years + 1):
            for month in range(1, 13):
                # Contributions mensuelles
                total_amount += self.monthly_contribution
                total_invested += self.monthly_contribution

                # Évolution du marché
                growth_factor = (1 + self.market_growth) ** (1 / 12)
                market_change = total_amount * (growth_factor - 1)
                total_amount += market_change

                # Frais de gestion trimestriels
                fees = (total_amount * (self.management_fee / 100)) / 4 if month % 3 == 0 else 0
                total_amount -= fees
                total_fees += fees
                total_fees_cumulative += fees  # Ajout des frais au cumul

                # Dividendes versés en décembre
                dividends = total_amount * self.dividend_yield if month == 12 else 0
                total_amount += dividends
                total_dividends += dividends

                # Stocker les valeurs dans le DataFrame
                self.data.loc[f"{self.months[month-1]} {self.current_year + year - 1}"] = [
                    round(total_amount, 2),
                    round(total_invested, 2),
                    round(total_amount - total_invested, 2),
                    round(total_fees, 2),
                    round(total_fees_cumulative, 2),  # Nouvelle colonne ajoutée
                    round(market_change, 2),
                    round(dividends, 2),
                    round(self.monthly_contribution, 2)
                ]

    def get_data(self):
        return self.data

# Exemple d'utilisation
assurance_vie = InvestmentAccount(
    name="Assurance Vie",
    start_amount=5000,
    management_fee=0.75,
    market_growth=0.10,
    dividend_yield=0.0159,
    monthly_contribution=600
)

assurance_vie.simulate(years=30)  # Simulation sur 30 ans

frais_courtages = 2

pea = InvestmentAccount(
    name="PEA",
    start_amount=5000,
    management_fee=0,
    market_growth=0.10,
    dividend_yield=0.0159,
    monthly_contribution=600-frais_courtages  # -2€ de frais de courtage
)

pea.simulate(years=30)  # Simulation sur 30 ans


# print(assurance_vie.get_data())

print(assurance_vie.get_data().iloc[-12:]["Somme totale"])

# janvier_data = assurance_vie.get_data().filter(like="Mois 1", axis=0)
# print(janvier_data)

# # Exemple d'utilisation
# save_to_excel(assurance_vie, "investissement2.xlsx")

# plot_investment(assurance_vie, "Somme totale", "Bénéfices", "Somme totale investie","Montants totaux de frais de gestion")

plot_multiple_accounts(assurance_vie, pea, column="Somme totale")