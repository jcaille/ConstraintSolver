from ortools.sat.python import cp_model
from read_excel import DayShift, read_configuration

configuration = read_configuration()


def var_name(agent, day, shift, posting):
    return f"{agent} | {day} | {shift} | {posting}"


model = cp_model.CpModel()
solver = cp_model.CpSolver()

AGENTS = [a.name for a in configuration.agents]
DAYS = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI"]
SHIFTS = ["MATIN", "APRES-MIDI"]
POSTINGS = ["FREE", "ECHO", "IRM", "TDM"]
VALUES = {"FREE": 0, "ECHO": 1, "IRM": 100, "TDM": 10000}

variables = {}


def get(agent, day, shift, posting):
    return variables[var_name(agent, day, shift, posting)]


print("Creating all variables")
for day in DAYS:
    for shift in SHIFTS:
        for posting in POSTINGS:
            for agent in AGENTS:
                name = var_name(agent, day, shift, posting)
                variables[name] = model.NewBoolVar(name)

"""
Humans can only be on one posting at any given time
"""
print("Creating 'humans can only be in one place' rule")
for agent in AGENTS:
    for day in DAYS:
        for shift in SHIFTS:
            model.Add(
                get(agent, day, shift, "FREE")
                + get(agent, day, shift, "ECHO")
                + get(agent, day, shift, "IRM")
                + get(agent, day, shift, "TDM")
                == 1
            )


"""
Pour chaque ligne n, la somme des coefficients est fixée, généralement différente d’une ligne à l’autre.
"""
print("Creating 'agents have a fixed number of slots' rule")
for agent in configuration.agents:
    values = [
        sum(get(agent.name, day, shift, "ECHO") for day in DAYS for shift in SHIFTS),
        sum(get(agent.name, day, shift, "IRM") for day in DAYS for shift in SHIFTS),
        sum(get(agent.name, day, shift, "TDM") for day in DAYS for shift in SHIFTS),
    ]
    targets = [agent.echo_capacity, agent.irm_capacity, agent.tdm_capacity]
    for i in range(0, len(values)):
        if configuration.vacation_maximum:
            model.Add(values[i] <= targets[i])
        else:
            model.Add(values[i] == targets[i])


"""
Certains radiologues ont des jours OFF réservés
"""
for agent in configuration.agents:
    for day_off in agent.off:
        model.Add(get(agent.name, day_off.day, day_off.shift, "FREE") == 1)

"""
Pour chaque colonne p la somme des coefficients est fixée, généralement différente d’une ligne à l’autre.
"""
print("Creating 'days have required demand' rules")
for day in DAYS:
    for shift in SHIFTS:
        dayshift = DayShift(day, shift)
        demand = configuration.demands[dayshift]

        values = [
            sum(get(agent, day, shift, "ECHO") for agent in AGENTS),
            sum(get(agent, day, shift, "IRM") for agent in AGENTS),
            sum(get(agent, day, shift, "TDM") for agent in AGENTS),
        ]
        targets = [demand.echo, demand.irm, demand.tdm]
        for i in range(0, len(values)):
            if configuration.staffing_minimum:
                model.Add(values[i] >= targets[i])
            else:
                model.Add(values[i] == targets[i])

"""
Certaines associations de postes au cours de la même journée ne sont pas autorisées.

- Meme poste le matin et l'apres midi
"""

for agent in AGENTS:
    for day in DAYS:
        # Not the same posting on all shifts
        for posting in POSTINGS:
            if posting == "FREE":
                continue
            model.AddForbiddenAssignments(
                [get(agent, day, shift, posting) for shift in SHIFTS],
                [[1 for _ in SHIFTS]],
            )

        # Not allowed on TDM and IRM in the same day
        model.AddForbiddenAssignments(
            [get(agent, day, "MATIN", "IRM"), get(agent, day, "APRES-MIDI", "TDM")],
            [[1, 1]],
        )
        model.AddForbiddenAssignments(
            [get(agent, day, "MATIN", "TDM"), get(agent, day, "APRES-MIDI", "IRM")],
            [[1, 1]],
        )
        model.AddForbiddenAssignments(
            [get(agent, day, "MATIN", "TDM"), get(agent, day, "APRES-MIDI", "ECHO")],
            [[1, 1]],
        )


status = solver.Solve(model)
if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
    print("================= Solution =================")
    print(f"Solved in {solver.WallTime():.2f} milliseconds")

    print("======== PLANNING PAR RADIOLOGUE ============")
    for agent in AGENTS:
        print(f"\n{agent:^50}")
        for day in DAYS:
            for shift in SHIFTS:
                for posting in POSTINGS:
                    if posting == "FREE":
                        continue
                    if solver.Value(get(agent, day, shift, posting)):
                        print(
                            f"{agent:<10} - Posté(e) en {posting:<4} le {day:<5} {shift}"
                        )

    print("\n\n")
    print("======== PLANNING PAR POSTE ============")
    for day in DAYS:
        for shift in SHIFTS:
            for posting in POSTINGS:
                if posting == "FREE":
                    continue
                posted = [
                    a for a in AGENTS if solver.Value(get(a, day, shift, posting))
                ]
                print(f"{posting:<4} le {day:<8} {shift:<10} : {', '.join(posted)}")

    print("================= MERCI !================")
else:
    print("The solver could not find a solution.")
