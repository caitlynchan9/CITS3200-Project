<template>
  <div>
    <div class="d-flex mb-2 display-3 font-weight-medium indigo--text text--darken-2">Publish</div>
    <div
      class="d-flex mb-10 display-1 grey--text text--darken-1"
    >Select publish to create template or return to previous page</div>

    <v-card flat color="rgb(255, 0, 0, 0)">
        <v-card-actions>
    <v-btn x-large color="indigo darken-3" class="white--text" href="/Competencies">
            <v-icon>mdi-account-arrow-left</v-icon>Back to previous page
            </v-btn>
        <v-spacer></v-spacer>
        
        <v-btn x-large color="indigo darken-3" class="white--text">
            <v-icon>mdi-publish</v-icon>Publish Interview Guide
            </v-btn>
</v-card-actions>
    </v-card>

    <VueDocPreview value="newteach.pbworks.com%2Ff%2Fele%2Bnewsletter.docx" type="office" />
  </div>
</template>

<script>
import VueDocPreview from "vue-doc-preview";
export default {
  components: {
    VueDocPreview
  }
};
</script>