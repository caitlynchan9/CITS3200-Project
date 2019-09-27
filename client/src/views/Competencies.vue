<template>
  <div>
    <div class="d-flex mb-2 display-3 font-weight-medium indigo--text text--darken-2" >Competencies</div>
    <div class="d-flex mb-10 ml-2 display-1 grey--text text--darken-1">Select the actions for each desired competency</div>

    <v-layout column wrap>
      <v-flex xs12 md6></v-flex>
      <v-expansion-panels v-model="panel" multiple>
        <v-expansion-panel>
          <v-expansion-panel-header>Competency 1</v-expansion-panel-header>
          <v-expansion-panel-content>
            <v-checkbox label="Disabled"></v-checkbox>
          </v-expansion-panel-content>
        </v-expansion-panel>

        <v-expansion-panel>
          <v-expansion-panel-header>Competency 2</v-expansion-panel-header>
          <v-expansion-panel-content>
            <v-checkbox label="Disabled"></v-checkbox>
          </v-expansion-panel-content>
        </v-expansion-panel>

        <v-expansion-panel>
          <v-expansion-panel-header>Competency 3</v-expansion-panel-header>
          <v-expansion-panel-content>
            <v-checkbox label="Disabled"></v-checkbox>
          </v-expansion-panel-content>
        </v-expansion-panel>
      </v-expansion-panels>
    </v-layout>
  </div>
</template>

<script>
export default {
  data: () => ({
    panel: [0, 1]
  })
};
</script>