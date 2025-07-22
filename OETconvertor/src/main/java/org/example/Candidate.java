package org.example;
public class Candidate {

    private String candidateNumber;
    private String firstName;
    private String lastName;
    private String profession;

    // constructor
    public Candidate(String candidateNumber, String firstName, String lastName, String profession) {
        this.candidateNumber = candidateNumber;
        this.firstName = firstName;
        this.lastName = lastName;
        this.profession = profession;
    }

    // getters
    public String getCandidateNumber() { return candidateNumber; }
    public String getFirstName() { return firstName; }
    public String getLastName() { return lastName; }
    public String getProfession() { return profession; }

    // setters
    public void setFirstName(String firstName) { this.firstName = firstName; }
    public void setLastName(String lastName) { this.lastName = lastName; }
    public void setProfession(String profession) { this.profession = profession; }
}