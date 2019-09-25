<template>
  <v-card color="grey lighten-4" flat height="200px" tile>
    <v-app-bar app extended clipped-left class="indigo darken-4">
      <v-toolbar-title class="white--text">
        <img class="mt-12" :src="require('@/assets/uwa-crest.png')" height="80" />
        Talent Acquisition
      </v-toolbar-title>


      <div class="flex-grow-1"></div>

      <v-btn icon>
        <v-icon>mdi-magnify</v-icon>
      </v-btn>

      <v-btn icon>
        <v-icon>mdi-heart</v-icon>
      </v-btn>

      <v-btn icon>
        <v-icon>mdi-dots-vertical</v-icon>
      </v-btn>
    </v-app-bar>

    <v-navigation-drawer app class="indigo darken-1" clipped dark permanent>
      <v-list>
        <v-list-item v-for="item in items" :key="item.title" 
        @click="":to="item.route">
          <v-list-item-icon>
            <v-icon>{{ item.icon }}</v-icon>
          </v-list-item-icon>

          <v-list-item-content>
            <v-list-item-title>{{ item.title }}</v-list-item-title>
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
      items: [
        { title: "Home", icon: "mdi-home", route: "/" },
        { title: "Competencies", icon: "mdi-clipboard-list-outline", route: "/competencies" },
        { title: "Search", icon: "mdi-magnify", route: "/search" }
      ]
    };
  }
};
</script>