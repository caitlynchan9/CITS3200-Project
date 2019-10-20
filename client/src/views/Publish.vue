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
    <div style="height: 900px;">
      <iframe
      width="50%"
      height="100%"
        src="https://docs.google.com/document/d/e/2PACX-1vTFa_TH4xqMT9wAh13xPb24PluPqE8jXkA4jMNyFq6BtjFiyLngGe9KJP3kX9QE5w/pub?embedded=true"
      ></iframe>
   </div>
  </div>
</template>

<script>
</script>