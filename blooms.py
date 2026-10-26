import re


BLOOM_VERBS = {
    "Remember": [
        "define","identify","list","name","recall","state","label","mention"
    ],

    "Understand": [
        "describe","explain","summarize","classify","interpret","discuss"
    ],

    "Apply": [
        "apply","use","implement","execute","solve","demonstrate","perform"
    ],

    "Analyze": [
        "analyze","compare","differentiate","categorize","examine","investigate"
    ],

    "Evaluate": [
        "evaluate","justify","recommend","assess","critique","validate"
    ],

    "Create": [
        "design","develop","construct","create","build","formulate"
    ]
}

def detect_bloom(statement):

    statement = str(statement).lower()

    for level, verbs in BLOOM_VERBS.items():

        for verb in verbs:

            if re.search(r"\b" + verb + r"\b", statement):

                return level

    return "Understand"