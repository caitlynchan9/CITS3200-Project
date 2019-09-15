<template>
  <v-card>
    <v-container class="pa-2" fluid>
      <v-row>
        <v-col>
          <v-card class="indigo darken-4" dark>
            <v-card-text class="white--text">
              <div class="headline mb-2">LEADING & DECIDING</div>
              <v-container fluid>
                <v-checkbox
                  v-model="ex4"
                  label="1. Decigin and Initiating Action"
                  color="white"
                  value="red"
                  hide-details
                ></v-checkbox>
              </v-container>
            </v-card-text>
          </v-card>
        </v-col>

        <v-col>
          <v-card class="indigo darken-4" dark>
            <v-card-text class="white--text">
              <div class="headline mb-2">SUPPORTING & CO-OPERATING</div>
              <v-container fluid>
                <v-checkbox
                  v-model="ex4"
                  label="2. Adhering to Principles and Values"
                  color="white"
                  value="red"
                  hide-details
                ></v-checkbox>
              </v-container>
            </v-card-text>
          </v-card>
        </v-col>

        <v-col>
          <v-card class="indigo darken-4" dark>
            <v-card-text class="white--text">
              <div class="headline mb-2">LEADING & DECIDING</div>
              <v-container fluid>
                <v-checkbox
                  v-model="ex4"
                  label="3. Relating and Networking"
                  color="white"
                  value="red"
                  hide-details
                ></v-checkbox>
              </v-container>
            </v-card-text>
          </v-card>
        </v-col>
      </v-row>
    </v-container>
  </v-card>
</template>

<script>
export default {
  data: () => ({})
};
</script>