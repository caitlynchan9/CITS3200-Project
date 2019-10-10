<template>
  <div>
    <div class="d-flex mb-2 display-3 font-weight-medium indigo--text text--darken-2">Competencies</div>
    <div
      class="d-flex mb-10 ml-2 display-1 grey--text text--darken-1"
    >Select the actions for each desired competency</div>

    <v-layout column wrap>
      <v-flex xs12 md6></v-flex>
      <v-expansion-panels v-model="panel" multiple>
        <v-expansion-panel v-for="(comp, index) in competencies" :key="index">
          <v-expansion-panel-header>{{ comp.id }}</v-expansion-panel-header>
          <v-expansion-panel-content v-for="question in comp.questions" :key="q">
            <v-layout row wrap>
              <v-checkbox :label="question"></v-checkbox>
            </v-layout>
          </v-expansion-panel-content>
        </v-expansion-panel>

        <v-flex xs12 offset-xs9>
          <div class="my-2">
            <v-tooltip left>
              <template v-slot:activator="{ on }">
                <v-btn x-large color="indigo darken-3" dark v-on="on" href="/Publish">Create Interview Guide</v-btn>
              </template>
              <span>Create an Interview Guide from the selected competencies</span>
            </v-tooltip>
          </div>
        </v-flex>
      </v-expansion-panels>
    </v-layout>
  </div>
</template>



<script>
export default {
  data: () => ({
    panel: [0, 1],
    competencies: [
      {
        id: "1. Deciding and Initiating Action / Taking Responsibility",
        questions: [
          "Give me and example of a time when you did not receive any guidance from your supervisor and had to decide for yourself what to do at work",
          "Tell me about a time when you asked someone else for help with your job/task.",
          "Give me an example of a situation in which you had to make a difficult or challenging work-related decision"
        ]
      },
      {
        id: "2. Working with People / Supporting Others",
        questions: [
          "Tell me about the last time you had to help someone at work with his or her job/tasks.",
          "Tell me about the last time you worked as part of a team/worked with someone else on a task.",
          "What qualities do you think are necessary to make teams effective."
        ]
      },
      {
        id: "3. Adhering to Principles and Values / Acting with Integrity",
        questions: ["What sort of ethics or values do you need to uphold in your work?", "Describe a time when you had to demonstrate high levels of integrity in your work and actions.", "Give an example of an occasion when you have been particularly careful to treat others fairly."]
      },
      {
        id: "4. Relating and Networking / Relating to Others",
        questions: ["Tell me about the last time that a new employee joined your work team.", "Give me an example where there was a disagreement amongst colleagues in your work team.", "Tell me of a time when you had to make contact with another team or department in order to get something done that was to the benefit of the team or yourself."]
      },
      {
        id: "5. Presenting and Communicating Information / Oral Communication",
        questions: ["How did you tailor it to suit the audience?", "Describe a time when you won an argument or debate.", "In what situations do you find it difficult to get your point across?"]
      },
      {
        id: "6. Writing and Reporting / Written Communication",
        questions: ["What kind of correspondence have you had to prepare in the past?", "Tell me about your experience of preparing reports.", "Give me an example of a time you modified a standard letter or briefing."]
      }
    ]
  })
};
</script>