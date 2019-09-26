<template>
  <div>
    <div class="d-flex">Competencies</div>

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
  </div>
</template>

<script>
export default {
  data: () => ({
    panel: [0, 1]
  })
};
</script>