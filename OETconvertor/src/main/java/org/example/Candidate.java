package org.example;

public class Candidate {
    private final String candidateNumber;
    private final String firstName;
    private final String lastName;
    private final String profession;

    public Candidate(String candidateNumber, String firstName, String lastName, String profession) {
        this.candidateNumber = candidateNumber;
        this.firstName = firstName;
        this.lastName = lastName;
        this.profession = profession;
    }

    public String getCandidateNumber() {
        return candidateNumber;
    }

    public String getFirstName() {
        return firstName;
    }

    public String getLastName() {
        return lastName;
    }

    public String getProfession() {
        return profession;
    }


    @Override
    public String toString() {
        return String.format("Candidate: %s %s (ID: %s, Profession: %s)",
                firstName, lastName, candidateNumber, profession);
    }
}