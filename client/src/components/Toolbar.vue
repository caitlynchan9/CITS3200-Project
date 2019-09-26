<template>
  <v-card color="grey lighten-4" flat height="200px" tile>
    <v-app-bar app clipped-left class="indigo darken-4">
      <v-toolbar-title class="white--text">
        <!-- <img class="mt-12" :src="require('@/assets/uwa-crest.png')" height="80" /> -->
        Talent Acquisition
      </v-toolbar-title>

      <div class="flex-grow-1"></div>

      <v-toolbar-items>
        <v-btn text>
            <v-icon>mdi-home</v-icon>
            Home</v-btn>
        <v-btn text>
            <v-icon>mdi-account-supervisor-outline</v-icon>
            Competencies</v-btn>
        <v-btn text>
            <v-icon>mdi-magnify</v-icon>
            Search</v-btn>
      </v-toolbar-items>

    </v-app-bar>

    <v-navigation-drawer app class="indigo darken-1" clipped dark permanent>
      <v-list>
        <v-list-item v-for="comp in competencies" :key="comp.title" @click :to="comp.route">
          <v-list-item-icon>
            <v-icon>{{ comp.icon }}</v-icon>
          </v-list-item-icon>

          <v-list-item-content>
            <v-list-item-title>{{ comp.id }}</v-list-item-title>
          </v-list-item-content>
        </v-list-item>
      </v-list>

      <template v-slot:append>
        <div class="pa-2">
          <v-btn block>Logout</v-btn>
        </div>
      </template>
    </v-navigation-drawer>
  </v-card>
</template>


<script>
export default {
  data() {
    return {
      competencies: [
        { id: "Home", icon: "mdi-home", route: "/" },
        {
          id: "Competencies",
          icon: "mdi-clipboard-list-outline",
          route: "/competencies"
        },
        { id: "Search", icon: "mdi-magnify", route: "/search" }
      ]
    };
  }
};
</script>