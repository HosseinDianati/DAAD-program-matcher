"""
Modified Bavarian Formula Calculator for German University Admissions Equivalency.

Formula:
    N = 1 + 3 * (N_max - N_d) / (N_max - N_min)

Where:
    N     = Converted German grade (1.0 = best, 4.0 = passing threshold)
    N_max = Maximum achievable grade in the source country's grading system
    N_min = Minimum passing grade in the source country's grading system
    N_d   = Candidate's obtained grade
"""

def calculate_german_gpa(obtained: float, max_grade: float, min_passing: float) -> float:
    """
    Calculates the German GPA equivalent rounded to 2 decimal places.
    """
    if obtained > max_grade:
        raise ValueError(f"Obtained grade ({obtained}) exceeds maximum possible grade ({max_grade}).")
    if obtained < min_passing:
        return 5.0  # Fail in German grading system
    
    german_grade = 1.0 + 3.0 * ((max_grade - obtained) / (max_grade - min_passing))
    # German grades are usually bounded between 1.0 and 4.0
    return round(max(1.0, min(4.0, german_grade)), 2)


# Predefined grading scale standards for international applicants
COUNTRY_SCALES = {
    "iran_20": {"max": 20.0, "min": 10.0, "name": "Iran (0-20 scale)"},
    "iran_bachelor": {"max": 20.0, "min": 10.0, "name": "Iran Bachelor (passing grade 10)"},
    "iran_master": {"max": 20.0, "min": 12.0, "name": "Iran Master (passing grade 12)"},
    "us_4": {"max": 4.0, "min": 2.0, "name": "US / 4.0 scale"},
    "india_10": {"max": 10.0, "min": 4.0, "name": "India 10-point CGPA (passing grade 4.0)"},
    "india_percentage": {"max": 100.0, "min": 40.0, "name": "India Percentage (passing 40%)"},
    "turkey_4": {"max": 4.0, "min": 2.0, "name": "Turkey 4.0 scale"},
    "egypt_percentage": {"max": 100.0, "min": 50.0, "name": "Egypt Percentage (passing 50%)"},
}


def get_german_grade_classification(german_grade: float) -> str:
    """Returns the qualitative German academic classification."""
    if german_grade <= 1.5:
        return "Sehr Gut (Very Good / Excellent)"
    elif german_grade <= 2.5:
        return "Gut (Good)"
    elif german_grade <= 3.5:
        return "Befriedigend (Satisfactory)"
    elif german_grade <= 4.0:
        return "Ausreichend (Sufficient / Pass)"
    else:
        return "Nicht Ausreichend (Fail)"


if __name__ == "__main__":
    import sys
    # Test for Mehrshad Jalilikhouei (18.05 / 20.0)
    gpa = calculate_german_gpa(18.05, 20.0, 10.0)
    desc = get_german_grade_classification(gpa)
    print(f"Candidate: 18.05/20 -> German GPA: {gpa} ({desc})")
